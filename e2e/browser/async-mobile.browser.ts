import { test, expect, devices, type Locator, type Page, type TestInfo } from 'playwright/test';
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';

type Format = 'docx' | 'xlsx' | 'pptx';
const formats: Format[] = ['docx', 'xlsx', 'pptx'];
const root = resolve(import.meta.dirname, '../..');
const typed = 'MobileTouchProbe';
const suffix = 'Rotated';
const marker = typed + suffix;
const phone = { width: 412, height: 915 };
const tablet = { width: 1180, height: 820 };

test.setTimeout(30_000);

const chrome = {
  docx: { rail: 'formatting-bar', more: 'toolbar-more', menu: 'More actions' },
  xlsx: { rail: 'xlsx-formatting-toolbar', more: 'xlsx-toolbar-more', menu: 'More toolbar items' },
  pptx: { rail: 'pptx-formatting-toolbar', more: 'pptx-toolbar-more', menu: 'More' },
};

const toolbarGroups: Record<Format, [string, string[]][]> = {
  docx: [
    ['History', ['Undo', 'Redo']],
    ['Zoom', ['Zoom level']],
    ['Styles', ['Paragraph style']],
    ['Font', ['Font', 'Font size']],
    ['Text formatting', ['Bold', 'Italic', 'Underline', 'Strikethrough', 'Font Color', 'Text Highlight Color', 'Insert link']],
    ['Script', ['Superscript', 'Subscript']],
    ['Alignment', ['Alignment']],
    ['List formatting', ['Bullet List', 'Numbered List', 'Decrease Indent', 'Increase Indent', 'Line spacing']],
  ],
  xlsx: [
    ['History', ['Undo', 'Redo']],
    ['File', ['Save', 'Export PNG']],
    ['Print', ['Print']],
    ['Paint format', ['Paint format']],
    ['Zoom', ['Zoom']],
    ['Number formats', ['Format as currency', 'Format as percent', 'Decrease decimal places', 'Increase decimal places', 'Number format']],
    ['Font', ['Font family']],
    ['Font', ['Decrease font size', 'Font size', 'Increase font size']],
    ['Text formatting', ['Bold', 'Italic', 'Strikethrough']],
    ['Colors', ['Text color', 'Fill color']],
    ['Borders', ['Borders']],
    ['Merge cells', ['Merge cells']],
    ['Alignment', ['Horizontal alignment', 'Vertical alignment', 'Text wrapping']],
  ],
  pptx: [
    ['History', ['Undo', 'Redo']],
    ['File', ['Save', 'Export PNG']],
    ['Slides', ['New slide']],
    ['Zoom', ['Zoom']],
    ['Tools', ['Select', 'Text box', 'Insert image', 'Shape']],
    ['Font', ['Font family', 'Decrease font size', 'Font size', 'Increase font size']],
    ['Text formatting', ['Bold', 'Italic', 'Underline', 'Text color']],
    ['Alignment', ['Alignment']],
    ['Shape formatting', ['Fill color', 'Border color', 'Border width', 'Arrange']],
  ],
};

function demo(format: Format) {
  return resolve(root, 'apps/demo/public', format === 'xlsx' ? 'showcase.xlsx' : 'betteroffice-demo.' + format);
}

function canvas(page: Page, format: Format) {
  if (format === 'docx') return page.locator('[data-docx-worker-editor] .canvas-page canvas').first();
  if (format === 'xlsx') return page.getByTestId('xlsx-scroll').locator('canvas').first();
  return page.getByTestId('pptx-slide-canvas');
}

function textInput(page: Page, format: 'docx' | 'pptx') {
  return format === 'docx'
    ? page.getByRole('textbox', { name: 'Document input', exact: true })
    : page.getByRole('textbox', { name: 'Slide text', exact: true });
}

async function open(page: Page, format: Format, file: string = demo(format)) {
  await page.goto('/?format=' + format);
  await page.locator('input[type=file]').last().setInputFiles(file);
  await expect(canvas(page, format)).toBeVisible();
  await expect(page.getByRole('button', { name: 'Save', exact: true }).first()).toBeEnabled();
  await expect(page.locator('.notice, [data-testid$="-error"]')).toHaveCount(0);
  await page.evaluate(() => document.fonts.ready.then(() => undefined));
  if (format === 'docx') {
    await expect(textInput(page, format)).toBeAttached();
    await expect(page.locator('.canvas-page-mirror').first()).toContainText(/\S/);
  } else if (format === 'xlsx') {
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await expect(page.getByTestId('xlsx-selection')).toBeVisible();
  } else {
    await expect(page.getByTestId('pptx-worker-editor')).toBeVisible();
    await expect(page.locator('aside[aria-label="Slides"] button[aria-current="page"]')).toHaveCount(1);
  }
}

async function save(page: Page, format: Format, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().tap();
  const download = await downloading;
  expect(download.suggestedFilename()).toMatch(new RegExp('\\.' + format + '$'));
  const file = info.outputPath(label + '.' + format);
  await download.saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const names = Object.keys(zip.files).filter((name) => name.endsWith('.xml')).sort();
  const xml: Record<string, string> = Object.fromEntries(await Promise.all(names.map(async (name) =>
    [name, await zip.file(name)!.async('string')] as const
  )));
  const text = await page.evaluate((parts) => Object.fromEntries(Object.entries(parts).map(([name, xml]) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved XML: ' + name);
    return [name, Array.from(document.getElementsByTagNameNS('*', 't')).map((node) => node.textContent ?? '').join('')];
  })), xml);
  return { file, xml, text };
}

async function inViewport(locator: Locator) {
  await expect(locator).toBeVisible();
  await expect.poll(() => locator.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return rect.left >= -1 && rect.right <= window.innerWidth + 1;
  })).toBe(true);
}

async function noPageScroll(page: Page) {
  await expect.poll(() => page.evaluate(() => ({
    horizontal: document.documentElement.scrollWidth <= window.innerWidth + 1,
    x: window.scrollX,
    y: window.scrollY,
  }))).toEqual({ horizontal: true, x: 0, y: 0 });
}

async function fits(page: Page, format: Format) {
  await inViewport(page.locator('.editor-stage'));
  await noPageScroll(page);
  if (format === 'xlsx') {
    await inViewport(page.getByTestId('xlsx-scroll'));
    await inViewport(canvas(page, format));
    await inViewport(page.getByTestId('xlsx-name-box'));
  } else if (format === 'docx') {
    const scroller = page.locator('.docx-editor__scroll-container');
    await inViewport(scroller);
    await expect(scroller).toHaveCSS('overflow-x', 'auto');
  } else {
    const scroller = page.getByTestId('pptx-slide-canvas').locator('..');
    await expect.poll(() => scroller.evaluate((element) => element.scrollWidth <= element.clientWidth + 1)).toBe(true);
    await inViewport(canvas(page, format));
  }
}

async function moreMenu(page: Page, format: Format) {
  const config = chrome[format];
  await page.getByTestId(config.rail).getByTestId(config.more).tap();
  const menu = page.getByRole('menu', { name: config.menu, exact: true });
  await expect(menu).toBeVisible();
  return menu;
}

async function history(page: Page, format: Format, label: 'Undo' | 'Redo') {
  const rail = page.getByTestId(chrome[format].rail);
  const button = rail.getByRole('button', { name: label, exact: true });
  if (await button.isVisible()) {
    await expect(button).toBeEnabled();
    await button.tap();
  } else {
    const menu = await moreMenu(page, format);
    const item = menu.locator(`[data-label="${label}"]`);
    await expect(item).not.toHaveAttribute('aria-disabled', 'true');
    await item.tap();
    await expect(menu).toBeHidden();
  }
}

async function tapDocument(page: Page, format: Format, double = false) {
  let target: Locator;
  let x: number;
  let y: number;
  if (format === 'docx') {
    target = canvas(page, format);
    const box = (await target.boundingBox())!;
    // The demo's first word starts at the one-inch page margins on an 816 × 1056 page.
    x = box.x + box.width * 110 / 816;
    y = box.y + box.height * 114 / 1056;
  } else if (format === 'xlsx') {
    target = page.getByTestId('xlsx-selection');
    const box = (await target.boundingBox())!;
    x = box.x + Math.min(box.width / 2, 40);
    y = box.y + box.height / 2;
  } else {
    target = canvas(page, format);
    const box = (await target.boundingBox())!;
    // Inside "Office", clear of the title's six-screen-pixel shape border.
    x = box.x + box.width * 0.14;
    y = box.y + box.height * 0.24;
  }
  const point = await page.evaluate(({ x, y }) => {
    const viewport = window.visualViewport;
    return {
      x: (x - (viewport?.offsetLeft ?? 0)) * (viewport?.scale ?? 1),
      y: (y - (viewport?.offsetTop ?? 0)) * (viewport?.scale ?? 1),
    };
  }, { x, y });
  await page.touchscreen.tap(point.x, point.y);
  if (double) await page.touchscreen.tap(point.x, point.y);
  if (format === 'xlsx') {
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await expect(page.getByTestId('xlsx-selection')).toBeVisible();
    await expect(page.getByTestId(double ? 'xlsx-cell-editor' : 'xlsx-scroll')).toBeFocused();
  } else {
    await expect(textInput(page, format)).toBeFocused();
    if (double) {
      const selected = format === 'docx'
        ? page.locator('.oox-canvas-live-region')
        : page.getByTestId('pptx-worker-editor').getByRole('status').filter({ hasText: 'Text selected' });
      await expect(selected).toContainText(/selected/i);
    } else if (format === 'pptx') {
      await expect(page.getByTestId('pptx-text-edit-outline')).toBeVisible();
      await expect(page.getByTestId('pptx-caret')).toBeAttached();
    }
  }
}

async function visibleEdit(page: Page, format: Format) {
  if (format === 'docx') {
    await expect(page.locator('.canvas-page-mirror').first()).toContainText(marker);
  } else if (format === 'xlsx') {
    await expect(page.getByRole('gridcell').filter({ hasText: marker })).toHaveCount(1);
  } else {
    await expect(page.getByTestId('pptx-text-edit-outline')).toBeVisible();
    await expect(page.getByTestId('pptx-caret')).toBeAttached();
  }
}

async function cellValue(page: Page, xml: Record<string, string>) {
  return page.evaluate((parts) => {
    const parse = (xml: string) => new DOMParser().parseFromString(xml, 'application/xml');
    const worksheet = parse(parts['xl/worksheets/sheet1.xml']);
    const cell = Array.from(worksheet.getElementsByTagNameNS('*', 'c')).find((cell) => cell.getAttribute('r') === 'A1');
    if (!cell) throw new Error('Saved worksheet has no A1 cell');
    const value = cell.getElementsByTagNameNS('*', 'v')[0]?.textContent ?? '';
    if (cell.getAttribute('t') === 's') {
      const strings = parse(parts['xl/sharedStrings.xml']).getElementsByTagNameNS('*', 'si');
      return strings[Number(value)]?.textContent ?? '';
    }
    return cell.getAttribute('t') === 'inlineStr'
      ? Array.from(cell.getElementsByTagNameNS('*', 't')).map((node) => node.textContent ?? '').join('')
      : value;
  }, xml);
}

const slidePixels = (page: Page) => canvas(page, 'pptx').screenshot({
  style: '[data-testid="pptx-text-edit-outline"], [data-testid="pptx-caret-clip"], [data-testid="pptx-shape-selection"] { visibility: hidden !important; }',
});

async function editRotateSave(page: Page, format: Format, info: TestInfo, double: boolean) {
  const before = await save(page, format, info, 'before');
  if (format === 'pptx') {
    const box = (await canvas(page, format).boundingBox())!;
    await page.touchscreen.tap(box.x + box.width * 0.009375, box.y + box.height * 0.5);
    await expect(page.getByTestId('pptx-shape-selection')).toBeVisible();
  }
  await tapDocument(page, format, double);
  if (format === 'xlsx') {
    if (double) await page.keyboard.press('ControlOrMeta+A');
    await page.keyboard.type(typed);
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveValue(typed);
  } else {
    await page.keyboard.insertText(typed);
    if (format === 'docx') await expect(page.locator('.canvas-page-mirror').first()).toContainText(typed);
  }
  const viewport = page.viewportSize()!;
  const previousWidth = (await canvas(page, format).boundingBox())!.width;
  await page.setViewportSize({ width: viewport.height, height: viewport.width });
  await expect(canvas(page, format)).toBeVisible();
  if (format !== 'docx') {
    await expect.poll(async () => (await canvas(page, format).boundingBox())!.width).not.toBe(previousWidth);
  }
  await noPageScroll(page);
  const beforePixels = format === 'pptx' ? await slidePixels(page) : null;
  if (format === 'xlsx') {
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await expect(page.getByTestId('xlsx-cell-editor')).toBeFocused();
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveValue(typed);
    await page.keyboard.type(suffix);
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveValue(marker);
    await page.keyboard.press('Enter');
  } else {
    await expect(textInput(page, format)).toBeFocused();
    await page.keyboard.insertText(suffix);
  }
  await visibleEdit(page, format);
  const edited = await save(page, format, info, 'edited');
  const part = format === 'docx' ? 'word/document.xml' : 'ppt/slides/slide1.xml';
  if (format === 'xlsx') {
    expect(await cellValue(page, edited.xml)).toBe(marker);
  } else {
    expect(edited.text[part]).toContain(marker);
    if (double) {
      const word = format === 'docx' ? 'Welcome' : 'Office';
      expect(edited.text[part]).toBe(before.text[part].replace(word, marker));
    }
  }
  if (beforePixels) {
    await expect.poll(async () => beforePixels.equals(await slidePixels(page))).toBe(false);
  }
  const editedPixels = beforePixels ? await slidePixels(page) : null;
  await history(page, format, 'Undo');
  if (format === 'docx') await expect(page.locator('.canvas-page-mirror').first()).not.toContainText(marker);
  if (format === 'xlsx') await expect(page.getByRole('gridcell').filter({ hasText: marker })).toHaveCount(0);
  const undone = await save(page, format, info, 'undone');
  const beforeRotation = { ...edited.text, [part]: edited.text[part]?.replace(marker, typed) };
  const splitHistory = format !== 'xlsx' && undone.text[part] !== before.text[part];
  if (splitHistory) {
    expect(undone.text).toEqual(beforeRotation);
    await history(page, format, 'Undo');
    expect((await save(page, format, info, 'undone-before-rotation')).text).toEqual(before.text);
  } else {
    expect(undone.text).toEqual(before.text);
  }
  if (splitHistory) {
    await history(page, format, 'Redo');
    expect((await save(page, format, info, 'redone-before-rotation')).text).toEqual(beforeRotation);
  }
  await history(page, format, 'Redo');
  if (editedPixels) {
    await expect.poll(async () => editedPixels.equals(await slidePixels(page))).toBe(true);
  } else {
    await visibleEdit(page, format);
  }
  const redone = await save(page, format, info, 'redone');
  expect(redone.text).toEqual(edited.text);
  await open(page, format, redone.file);
  if (format === 'pptx') {
    await tapDocument(page, format);
    const reopenedPixels = await slidePixels(page);
    expect(beforePixels!.equals(reopenedPixels)).toBe(false);
  } else {
    await visibleEdit(page, format);
  }
  const reopened = await save(page, format, info, 'reopened');
  expect(reopened.text).toEqual(redone.text);
  if (format === 'xlsx') expect(await cellValue(page, reopened.xml)).toBe(marker);
}

async function swipe(page: Page, scroller: Locator) {
  const box = (await scroller.boundingBox())!;
  const x = Math.min(box.x + box.width * 0.65, page.viewportSize()!.width - 12);
  const start = box.y + box.height * 0.8;
  const end = box.y + box.height * 0.2;
  await scroller.evaluate((element) => {
    element.setAttribute('data-mobile-scroll-ended', 'false');
    element.setAttribute('data-mobile-touch-ended', 'false');
    const ended = () => {
      if (element.getAttribute('data-mobile-touch-ended') !== 'true') return;
      element.setAttribute('data-mobile-scroll-ended', 'true');
      element.removeEventListener('scrollend', ended);
    };
    element.addEventListener('scrollend', ended);
  });
  const cdp = await page.context().newCDPSession(page);
  try {
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x, y: start, id: 1 }] });
    for (let step = 1; step <= 12; step++) {
      await page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())));
      await cdp.send('Input.dispatchTouchEvent', {
        type: 'touchMove', touchPoints: [{ x, y: start + (end - start) * step / 12, id: 1 }],
      });
    }
    await scroller.evaluate((element) => element.setAttribute('data-mobile-touch-ended', 'true'));
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
  } finally {
    await cdp.detach();
  }
}

async function scrollWorkbook(info: TestInfo) {
  const zip = await JSZip.loadAsync(await readFile(demo('xlsx')));
  const worksheet = (await zip.file('xl/worksheets/sheet1.xml')!.async('string'))
    .replace(/<dimension\b[^>]*\/>/, '<dimension ref="A1:F200"/>')
    .replace('</sheetData>', '<row r="200"><c r="A200" t="inlineStr"><is><t>Scroll target</t></is></c></row></sheetData>');
  zip.file('xl/worksheets/sheet1.xml', worksheet);
  const file = info.outputPath('scroll-workbook.xlsx');
  await writeFile(file, await zip.generateAsync({ type: 'nodebuffer' }));
  return file;
}

async function scrollDeck(info: TestInfo) {
  const zip = await JSZip.loadAsync(await readFile(demo('pptx')));
  let presentation = await zip.file('ppt/presentation.xml')!.async('string');
  let relationships = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
  let contentTypes = await zip.file('[Content_Types].xml')!.async('string');
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
  const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
  for (let number = 4; number <= 12; number++) {
    zip.file(`ppt/slides/slide${number}.xml`, slide);
    zip.file(`ppt/slides/_rels/slide${number}.xml.rels`, rels);
    presentation = presentation.replace('</p:sldIdLst>', `<p:sldId id="${255 + number}" r:id="mobileSlide${number}"/></p:sldIdLst>`);
    relationships = relationships.replace('</Relationships>', `<Relationship Id="mobileSlide${number}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${number}.xml"/></Relationships>`);
    contentTypes = contentTypes.replace('</Types>', `<Override PartName="/ppt/slides/slide${number}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  }
  zip.file('ppt/presentation.xml', presentation);
  zip.file('ppt/_rels/presentation.xml.rels', relationships);
  zip.file('[Content_Types].xml', contentTypes);
  const file = info.outputPath('scroll-deck.pptx');
  await writeFile(file, await zip.generateAsync({ type: 'nodebuffer' }));
  return file;
}

test.describe('phone async editors', () => {
  test.use({
    viewport: phone,
    deviceScaleFactor: 2.625,
    isMobile: true,
    hasTouch: true,
    userAgent: devices['Pixel 7'].userAgent,
  });

  for (const format of formats) {
    test(`${format}: phone ${format === 'docx' ? 'viewport' : 'fit'} and save`, async ({ page }, info) => {
      await open(page, format);
      await fits(page, format);
      const saved = await save(page, format, info, 'phone');
      expect(Object.values(saved.text).join('')).not.toBe('');
      await open(page, format, saved.file);
      await fits(page, format);
      expect((await save(page, format, info, 'phone-reopened')).text).toEqual(saved.text);
    });

    test(`${format}: phone toolbar overflow`, async ({ page }, info) => {
      await open(page, format);
      await tapDocument(page, format);
      const before = await save(page, format, info, 'toolbar-before');
      const rail = page.getByTestId(chrome[format].rail);
      const menu = await moreMenu(page, format);
      await inViewport(menu);
      const groups = rail.locator('[data-toolbar-items] > [role="group"]');
      await expect(groups).toHaveCount(toolbarGroups[format].length);
      for (const [index, [group, labels]] of toolbarGroups[format].entries()) {
        const unit = groups.nth(index);
        await expect(unit).toHaveAttribute('aria-label', group);
        if (await unit.getAttribute('aria-hidden') === 'true') {
          for (const label of labels) await expect(menu.locator(`[data-label="${label}"]`)).toBeVisible();
        } else {
          await inViewport(unit);
        }
      }
      if (format === 'docx') {
        for (const label of ['Clear formatting', 'Toggle comments sidebar', 'Editing mode']) {
          await expect(menu.locator(`[data-label="${label}"]`)).toBeVisible();
        }
      }
      const font = format === 'docx' ? 'Font' : 'Font family';
      await menu.locator(`[data-label="${font}"]`).tap();
      const submenu = page.getByRole('menu', { name: font, exact: true });
      await inViewport(submenu);
      await expect(submenu.getByRole('menuitemradio').first()).toBeVisible();
      await page.getByTestId(chrome[format].more).tap();
      await expect(menu).toBeHidden();
      expect((await save(page, format, info, 'toolbar-after')).text).toEqual(before.text);
    });

    test(`${format}: phone tap typing and rotation`, async ({ page }, info) => {
      await open(page, format);
      await editRotateSave(page, format, info, false);
    });

    test(`${format}: phone touch scrolling`, async ({ page }, info) => {
      const file = format === 'pptx' ? await scrollDeck(info)
        : format === 'xlsx' ? await scrollWorkbook(info) : demo(format);
      await open(page, format, file);
      const before = await save(page, format, info, 'swipe-before');
      const scroller = format === 'docx'
        ? page.locator('.docx-editor__scroll-container')
        : format === 'xlsx'
          ? page.getByTestId('xlsx-scroll')
          : page.locator('aside[aria-label="Slides"]');
      await expect.poll(() => scroller.evaluate((element) => element.scrollHeight - element.clientHeight)).toBeGreaterThan(100);
      const top = await scroller.evaluate((element) => element.scrollTop);
      const pagePosition = await page.evaluate(() => ({ x: window.scrollX, y: window.scrollY }));
      await swipe(page, scroller);
      await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(top + 30);
      await expect(scroller).toHaveAttribute('data-mobile-scroll-ended', 'true');
      expect(await page.evaluate(() => ({ x: window.scrollX, y: window.scrollY }))).toEqual(pagePosition);
      expect((await save(page, format, info, 'swipe-after')).text).toEqual(before.text);
    });
  }
});

test.describe('tablet async editors', () => {
  test.use({
    viewport: tablet,
    deviceScaleFactor: 2,
    isMobile: true,
    hasTouch: true,
    userAgent: devices['iPad Pro 11 landscape'].userAgent,
  });

  for (const format of formats) {
    test(`${format}: tablet double tap and rotation`, async ({ page }, info) => {
      await open(page, format);
      await fits(page, format);
      await editRotateSave(page, format, info, true);
    });
  }
});
