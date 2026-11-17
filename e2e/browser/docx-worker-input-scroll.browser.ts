import { devices, expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const INPUT = '[data-testid="yrs-input"], .paged-editor__input';
const CARET = '[data-testid="caret"], [data-docx-pages-host] > div[aria-hidden="true"][style*="will-change: transform"]';
const SELECTION = '[data-testid^="selection-rect-"], [data-docx-pages-host] > div[aria-hidden="true"][style*="rgba(59, 130, 246"]';
const SCROLLER = '.docx-editor__scroll-container';
const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

declare global {
  interface Window { __docxScrollFocus: Array<boolean | undefined> }
}

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="link" Type="${r}/hyperlink" Target="https://example.com/scroll" TargetMode="External"/><Relationship Id="comments" Type="${r}/comments" Target="comments.xml"/></Relationships>`);
  zip.file('word/comments.xml', `<w:comments xmlns:w="${w}"><w:comment w:id="0" w:author="Reader" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Scroll comment</w:t></w:r></w:p></w:comment></w:comments>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>` +
    Array.from({ length: 4 }, (_, page) =>
      `<w:p>${page ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : ''}<w:r><w:t>Caret anchor ${page + 1}</w:t></w:r></w:p>` +
      '<w:p><w:hyperlink r:id="link"><w:r><w:t>Scroll link</w:t></w:r></w:hyperlink></w:p>' +
      (page === 0 ? '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Comment anchor</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>' : '') +
      Array.from({ length: 40 }, (_, row) => `<w:p><w:r><w:t>Line ${page + 1}.${row + 1}: selection follows the pointer across the page.</w:t></w:r></w:p>`).join('')
    ).join('') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function frames(page: Page, count = 4) {
  await page.evaluate(async (count) => {
    for (let frame = 0; frame < count; frame++) await new Promise<void>(resolve => requestAnimationFrame(() => resolve()));
  }, count);
}

async function open(page: Page, arm: string) {
  await page.goto('/?format=docx' + arm);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'input-scroll.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await documentBytes(),
  });
  const surface = page.locator('.canvas-page, .paged-editor__page').first();
  await expect(surface.locator('canvas').first()).toBeVisible({ timeout: 60_000 });
  await surface.click({ position: { x: 110, y: 105 } });
  const input = page.locator(INPUT).first();
  await expect(input).toBeFocused();
  await input.press('ControlOrMeta+Home');
  await expect(page.locator(CARET).first()).toBeAttached();
  await frames(page);
  return input;
}

async function scroll(page: Page, top: number) {
  await page.locator(SCROLLER).evaluate((element, top) => { element.scrollTop = top; }, top);
  await frames(page, 12);
  expect(await page.locator(SCROLLER).evaluate(element => element.scrollTop)).toBeCloseTo(top, 0);
}

async function rects(page: Page) {
  return page.evaluate(({ input, caret }) => {
    const a = document.querySelector(input)!.getBoundingClientRect();
    const b = document.querySelector(caret)!.getBoundingClientRect();
    return { input: { x: a.left, y: a.top, height: a.height }, caret: { x: b.left, y: b.top, height: b.height } };
  }, { input: INPUT, caret: CARET });
}

async function aligned(page: Page) {
  await expect.poll(async () => {
    const { input, caret } = await rects(page);
    return Math.max(Math.abs(input.x - caret.x), Math.abs(input.y - caret.y), Math.abs(input.height - caret.height));
  }).toBeLessThanOrEqual(2);
}

async function visibleCaret(page: Page) {
  await expect.poll(() => page.evaluate(({ caret, scroller }) => {
    const a = document.querySelector(caret)!.getBoundingClientRect();
    const b = document.querySelector(scroller)!.getBoundingClientRect();
    return a.top >= b.top && a.bottom <= b.bottom;
  }, { caret: CARET, scroller: SCROLLER })).toBe(true);
}

async function position(page: Page) {
  return page.locator(SCROLLER).evaluate(element => ({ top: element.scrollTop, left: element.scrollLeft, x: window.scrollX, y: window.scrollY }));
}

async function cursorAtLink(page: Page) {
  const link = page.locator('.canvas-page-mirror a[href="https://example.com/scroll"]').first();
  await expect(link).toBeAttached();
  const box = (await link.boundingBox())!;
  const point = { x: box.x + box.width / 2, y: box.y + box.height / 2 };
  await page.mouse.move(point.x, point.y);
  await frames(page);
  return page.evaluate(({ x, y }) => {
    const element = document.elementFromPoint(x, y)!;
    return { cursor: getComputedStyle(element).cursor, textarea: element instanceof HTMLTextAreaElement };
  }, point);
}

for (const [arm, label] of [['', 'worker'], ['&peerless=0', 'non-worker']] as const) {
  test(`docx input follows the caret through scrolling (${label})`, async ({ page }) => {
    test.skip(label === 'non-worker', 'The non-worker editor keeps its own input placement, which this change does not touch.');
    await open(page, arm);
    await aligned(page);
    await scroll(page, 80);
    await aligned(page);
    await scroll(page, 160);
    await aligned(page);
  });

  test(`docx composition starts at the scrolled caret and commits there (${label})`, async ({ page }) => {
    test.skip(label === 'non-worker', 'The non-worker editor keeps its own input placement, which this change does not touch.');
    const input = await open(page, arm);
    await input.press('ArrowRight');
    await frames(page);
    await scroll(page, 80);
    const caret = (await rects(page)).caret;
    const samples = await input.evaluate((element, caret) => {
      const textarea = element as HTMLTextAreaElement;
      const samples: Array<{ x: number; y: number; height: number }> = [];
      const sample = () => {
        const rect = textarea.getBoundingClientRect();
        samples.push({ x: rect.left - caret.x, y: rect.top - caret.y, height: rect.height - caret.height });
      };
      textarea.addEventListener('compositionstart', sample, { once: true, capture: true });
      textarea.addEventListener('compositionupdate', sample, { once: true, capture: true });
      textarea.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }));
      textarea.value = '語';
      textarea.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '語' }));
      textarea.dispatchEvent(new InputEvent('beforeinput', {
        bubbles: true, cancelable: true, inputType: 'insertCompositionText', data: '語', isComposing: true,
      }));
      textarea.dispatchEvent(new InputEvent('input', {
        bubbles: true, inputType: 'insertCompositionText', data: '語', isComposing: true,
      }));
      textarea.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '語' }));
      return samples;
    }, caret);
    expect(samples).toHaveLength(2);
    for (const sample of samples) expect(Math.max(Math.abs(sample.x), Math.abs(sample.y), Math.abs(sample.height))).toBeLessThanOrEqual(2);
    await expect(page.locator('.canvas-page-mirror').first()).toContainText('C語aret anchor 1');
    await expect(input).toHaveValue('');
    await aligned(page);
  });

  test(`docx typing reveals an offscreen caret once without a later jump (${label})`, async ({ page }) => {
    test.skip(label === 'non-worker', 'The non-worker editor keeps its own input placement, which this change does not touch.');
    const input = await open(page, arm);
    await scroll(page, 650);
    const offscreen = await page.evaluate(({ caret, scroller }) =>
      document.querySelector(caret)!.getBoundingClientRect().bottom < document.querySelector(scroller)!.getBoundingClientRect().top,
    { caret: CARET, scroller: SCROLLER });
    expect(offscreen).toBe(true);
    await page.keyboard.insertText('Typed');
    await expect(page.locator('.canvas-page-mirror').first()).toContainText('TypedCaret anchor 1');
    await visibleCaret(page);
    const revealed = await position(page);
    expect(revealed.top).toBeLessThan(650);
    await frames(page, 30);
    expect(await position(page)).toEqual(revealed);
    await scroll(page, 650);
    await frames(page, 30);
    expect((await position(page)).top).toBe(650);
    await expect(input).toBeFocused();
  });

  test(`docx selection drag extends beyond the viewport while scrolling (${label})`, async ({ page }) => {
    test.skip(label === 'non-worker', 'The non-worker editor keeps its own input placement, which this change does not touch.');
    await page.setViewportSize({ width: 1440, height: 600 });
    await open(page, arm);
    const caret = (await rects(page)).caret;
    const viewport = (await page.locator(SCROLLER).boundingBox())!;
    const bottom = viewport.y + viewport.height;
    await page.mouse.move(caret.x + 2, caret.y + caret.height / 2);
    await page.mouse.down();
    try {
      await page.mouse.move(caret.x + 120, bottom - 4, { steps: 20 });
      for (let step = 0; step < 50; step++) {
        await page.mouse.move(caret.x + 120 + step % 2, bottom + 12);
        await frames(page, 2);
      }
      await expect.poll(() => page.locator(SCROLLER).evaluate(element => element.scrollTop)).toBeGreaterThan(100);
    } finally { await page.mouse.up(); }
    await expect.poll(() => page.evaluate(({ selection, scroller }) => {
      const top = document.querySelector(scroller)!.scrollTop;
      return Math.max(0, ...Array.from(document.querySelectorAll(selection), element => element.getBoundingClientRect().bottom + top));
    }, { selection: SELECTION, scroller: SCROLLER })).toBeGreaterThan(bottom + 50);
    await expect(page.getByRole('button', { name: 'Add comment', exact: true })).toBeVisible();
  });

  test(`docx link cursor survives scrolling (${label})`, async ({ page }) => {
    await open(page, arm);
    const before = await cursorAtLink(page);
    expect(before.textarea).toBe(false);
    expect(before.cursor).not.toBe('');
    await page.mouse.move(4, 4);
    await scroll(page, 80);
    expect(await cursorAtLink(page)).toEqual(before);
  });

  test(`docx comment hover survives scrolling (${label})`, async ({ page }) => {
    await open(page, arm);
    const card = page.locator('.docx-comment-card', { hasText: 'Scroll comment' });
    if (!await card.isVisible()) await page.getByRole('button', { name: 'Toggle comments sidebar' }).click();
    await expect(card).toBeVisible();
    await card.click();
    const resolve = card.getByTitle('Resolve', { exact: true });
    await page.mouse.move(4, 4);
    const resting = await resolve.evaluate(element => getComputedStyle(element).backgroundColor);
    await resolve.hover();
    const hovered = await resolve.evaluate(element => getComputedStyle(element).backgroundColor);
    expect(hovered).not.toBe(resting);
    await page.mouse.move(4, 4);
    await scroll(page, 80);
    await resolve.hover();
    expect(await resolve.evaluate(element => getComputedStyle(element).backgroundColor)).toBe(hovered);
    await expect(card).toContainText('Scroll comment');
    await card.getByTitle('More options', { exact: true }).click();
    await expect(page.getByRole('menuitem', { name: 'Delete', exact: true })).toBeVisible();
  });

  test.describe(`docx mobile input scrolling (${label})`, () => {
    test.use({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 3,
      isMobile: true, hasTouch: true, userAgent: devices['iPhone 13'].userAgent });

    test('focusing the input preserves page and scroller positions', async ({ page }) => {
      await page.addInitScript(() => {
        window.__docxScrollFocus = [];
        const focus = HTMLTextAreaElement.prototype.focus;
        HTMLTextAreaElement.prototype.focus = function (this: HTMLTextAreaElement, options?: FocusOptions) {
          if (this.matches('[data-testid="yrs-input"], .paged-editor__input')) window.__docxScrollFocus.push(options?.preventScroll);
          focus.call(this, options);
        };
      });
      const input = await open(page, arm);
      await input.evaluate(element => (element as HTMLTextAreaElement).blur());
      await scroll(page, 650);
      const before = await position(page);
      await input.evaluate(element => (element as HTMLTextAreaElement).focus({ preventScroll: true }));
      await expect(input).toBeFocused();
      await frames(page, 30);
      expect(await position(page)).toEqual(before);
      const calls = await page.evaluate(() => window.__docxScrollFocus);
      expect(calls.length).toBeGreaterThan(1);
      expect(calls.every(value => value === true)).toBe(true);
    });
  });
}
