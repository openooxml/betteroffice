import { test, expect, type Locator, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { pinnedEvaluate } from './pinned-evaluate';

type Format = 'docx' | 'xlsx' | 'pptx';
type Saved = { file: string; parts: Record<string, string> };
const root = resolve(import.meta.dirname, '../..');
const pageCount = 16;
const marker = 'ZoomNavigationProbe';

async function fixture(format: Format, slides = 3) {
  const zip = await JSZip.loadAsync(await readFile(resolve(
    root,
    'apps/demo/public',
    format === 'xlsx' ? 'showcase.xlsx' : `betteroffice-demo.${format}`
  )));
  if (format === 'docx') {
    const paragraphs = Array.from({ length: pageCount }, (_, index) => {
      const number = String(index + 1).padStart(2, '0');
      const heading = `<w:p><w:pPr>${index ? '<w:pageBreakBefore/>' : ''}<w:spacing w:after="120"/></w:pPr><w:r><w:t>Navigation page ${number}</w:t></w:r></w:p>`;
      return heading + Array.from({ length: 31 }, (_, line) =>
        `<w:p><w:pPr><w:spacing w:after="80" w:line="280" w:lineRule="auto"/></w:pPr><w:r><w:t>Page ${number} line ${line + 1}. The quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>`
      ).join('');
    }).join('');
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  } else if (format === 'xlsx') {
    for (let sheet = 1; sheet <= 3; sheet++) {
      zip.file(`xl/worksheets/sheet${sheet}.xml`, `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:ZZ10000"/><sheetFormatPr defaultColWidth="10" defaultRowHeight="18"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Origin ${sheet}</t></is></c><c r="ZZ1" t="inlineStr"><is><t>Far column ${sheet}</t></is></c></row><row r="3"><c r="C3" t="inlineStr"><is><t>Hit target ${sheet}</t></is></c></row><row r="10000"><c r="A10000" t="inlineStr"><is><t>Far row ${sheet}</t></is></c><c r="ZZ10000" t="inlineStr"><is><t>Far corner ${sheet}</t></is></c></row></sheetData></worksheet>`);
    }
  } else {
    const presentation = await zip.file('ppt/presentation.xml')!.async('string');
    const relationships = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
    const slideRels = (await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string'))
      .replace(/<Relationship\b[^>]*Type="[^"]*\/notesSlide"[^>]*\/>/g, '');
    const ids: string[] = [];
    const rels: string[] = [];
    const types: string[] = [];
    for (let index = 1; index <= slides; index++) {
      ids.push(`<p:sldId id="${255 + index}" r:id="nav${index}"/>`);
      rels.push(`<Relationship Id="nav${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index}.xml"/>`);
      types.push(`<Override PartName="/ppt/slides/slide${index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
      zip.file(`ppt/slides/slide${index}.xml`, `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Navigation slide ${index}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Navigation target ${index}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="2438400" y="1371600"/><a:ext cx="3657600" cy="1714500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="315EFB"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Navigation slide ${index}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
      zip.file(`ppt/slides/_rels/slide${index}.xml.rels`, slideRels);
    }
    zip.file('ppt/presentation.xml', presentation.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${ids.join('')}</p:sldIdLst>`));
    zip.file('ppt/_rels/presentation.xml.rels', relationships
      .replace(/<Relationship\b[^>]*Type="[^"]*\/slide"[^>]*\/>/g, '')
      .replace('</Relationships>', `${rels.join('')}</Relationships>`));
    zip.file('[Content_Types].xml', contentTypes
      .replace(/<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, '')
      .replace('</Types>', `${types.join('')}</Types>`));
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, format: Format, file?: string, slides = 3) {
  await page.goto(`/?format=${format}`);
  await page.locator('input[type=file]').last().setInputFiles(file ?? {
    name: `navigation.${format}`,
    mimeType: 'application/octet-stream',
    buffer: await fixture(format, slides),
  });
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.locator('.notice, [data-testid$="-error"]')).toHaveCount(0);
  if (format === 'docx') {
    await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false');
    await expect(page.getByRole('status').filter({ hasText: /^\d+ of \d+$/ })).toHaveText(`1 of ${pageCount}`);
  } else if (format === 'xlsx') {
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Origin 1');
  } else {
    await expect(page.getByRole('complementary', { name: 'Slides', exact: true }).locator('button')).toHaveCount(slides);
  }
}

async function save(page: Page, format: Format, info: TestInfo, label: string): Promise<Saved> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const download = await downloading;
  const file = info.outputPath(`${label}.${format}`);
  await download.saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const parts: Record<string, string> = {};
  for (const part of Object.values(zip.files).filter((part) => !part.dir).sort((a, b) => a.name.localeCompare(b.name))) {
    parts[part.name] = format === 'docx' && part.name === 'docProps/core.xml'
      ? Buffer.from((await part.async('string')).replace(/(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified>)/g, '$1$2')).toString('base64')
      : await part.async('base64');
  }
  return { file, parts };
}

async function paragraphs(page: Page, saved: Saved) {
  const xml = Buffer.from(saved.parts['word/document.xml']!, 'base64').toString();
  return page.evaluate((xml) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved XML');
    return Array.from(document.getElementsByTagNameNS('*', 'p')).map((paragraph) =>
      Array.from(paragraph.getElementsByTagNameNS('*', 't')).map((text) => text.textContent ?? '').join('')
    );
  }, xml);
}

async function unchanged(page: Page, format: Format, info: TestInfo, original: Saved) {
  const navigated = await save(page, format, info, 'navigated');
  expect(navigated.parts).toEqual(original.parts);
  await open(page, format, navigated.file, format === 'pptx' ? Object.keys(original.parts).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length : 3);
  expect((await save(page, format, info, 'reopened')).parts).toEqual(original.parts);
}

function zoomControl(page: Page, format: Format) {
  return format === 'docx'
    ? page.getByRole('combobox', { name: /^Zoom:/ })
    : page.getByTestId(`${format}-zoom`);
}

async function preset(page: Page, format: Format, value: string) {
  const control = zoomControl(page, format);
  if (format === 'docx') await control.click();
  else await control.press('ArrowDown');
  await page.getByRole('option', { name: value, exact: true }).click();
  if (format === 'docx') await expect(control).toContainText(value);
  else await expect(control).toHaveValue(value);
}

async function typedZoom(page: Page, format: 'xlsx' | 'pptx', value: string) {
  const control = zoomControl(page, format);
  await control.fill(value);
  await control.press('Enter');
  await expect(control).toHaveValue(value);
}

async function width(locator: Locator) {
  return locator.evaluate((element) => element.getBoundingClientRect().width);
}

async function settled(locator: Locator) {
  let last = '';
  await expect.poll(async () => {
    const box = await locator.boundingBox();
    const key = JSON.stringify(box);
    const same = box !== null && key === last;
    last = key;
    return same;
  }).toBe(true);
  return (await locator.boundingBox())!;
}

async function fitted(canvas: Locator) {
  await expect.poll(() => canvas.evaluate((element) => {
    const stage = element.closest('[role="application"]')!;
    const expected = 1280 * Math.min(1, (stage.clientWidth - 40) / 1280, (stage.clientHeight - 40) / 720);
    return Math.abs(element.getBoundingClientRect().width - expected);
  })).toBeLessThan(1);
}

function stageOf(canvas: Locator): Locator {
  return canvas.locator('xpath=ancestor::*[@role="application"][1]');
}

async function inside(inner: Locator, outer: Locator) {
  const clip = await outer.boundingBox();
  expect(clip).not.toBeNull();
  await expect.poll(async () => {
    const box = await inner.boundingBox();
    return Boolean(box && box.width > 0 && box.height > 0 && box.x >= clip!.x - 1 && box.y >= clip!.y - 1 &&
      box.x + box.width <= clip!.x + clip!.width + 1 && box.y + box.height <= clip!.y + clip!.height + 1);
  }).toBe(true);
}

const docxInput = (page: Page) => page.locator('.paged-editor__input');
const docxStatus = (page: Page) => page.getByRole('status').filter({ hasText: /^\d+ of \d+$/ });
const docxScroller = (page: Page) => page.locator('.docx-editor__scroll-container');

async function ink(page: Page, clip: { x: number; y: number; width: number; height: number }) {
  const png = await page.screenshot({ clip: { x: clip.x, y: clip.y, width: clip.width, height: clip.height } });
  return pinnedEvaluate(page, async (base64) => {
    const bytes = Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
    const bitmap = await createImageBitmap(new Blob([bytes], { type: 'image/png' }));
    const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
    const context = canvas.getContext('2d')!;
    context.drawImage(bitmap, 0, 0);
    const pixels = context.getImageData(0, 0, bitmap.width, bitmap.height).data;
    bitmap.close();
    let count = 0;
    for (let index = 0; index < pixels.length; index += 4) {
      if (pixels[index]! + pixels[index + 1]! + pixels[index + 2]! < 300) count++;
    }
    return count;
  }, png.toString('base64'));
}

async function paintedDocx(page: Page) {
  await expect.poll(async () => {
    const clips = await page.locator('.canvas-page').evaluateAll((pages) => {
      let scroller = pages[0]?.parentElement ?? null;
      while (scroller && !/(auto|scroll)/.test(getComputedStyle(scroller).overflowY)) scroller = scroller.parentElement;
      if (!scroller) return [];
      const viewport = scroller.getBoundingClientRect();
      return pages.map((page) => {
        const box = page.getBoundingClientRect();
        const left = Math.max(box.left + 4, viewport.left);
        const right = Math.min(box.right - 4, viewport.right);
        const top = Math.max(box.top + 4, viewport.top + 40);
        const bottom = Math.min(box.bottom - 4, viewport.bottom);
        return { x: left, y: top, width: right - left, height: bottom - top, canvas: Boolean(page.querySelector('canvas')) };
      }).filter((clip) => clip.width > 100 && clip.height > 250);
    });
    if (!clips.length || clips.some((clip) => !clip.canvas)) return false;
    for (const clip of clips) if (await ink(page, clip) <= 50) return false;
    return true;
  }).toBe(true);
}

for (const format of ['docx', 'xlsx', 'pptx'] as const) {
  test(`${format}: zoom presets scale the painted surface without changing saved OOXML`, async ({ page }, info) => {
    await open(page, format);
    const original = await save(page, format, info, 'original');
    await preset(page, format, '100%');
    const surface = format === 'docx' ? page.locator('.canvas-page').first()
      : format === 'xlsx' ? page.getByTestId('xlsx-selection') : page.getByTestId('pptx-slide-canvas');
    await expect.poll(() => width(surface)).toBeGreaterThan(0);
    if (format === 'pptx') await expect.poll(() => width(surface)).toBeCloseTo(1280, 0);
    const initialWidth = (await settled(surface)).width;
    for (const value of ['50%', '150%', '100%']) {
      await preset(page, format, value);
      await expect.poll(() => width(surface)).toBeCloseTo(initialWidth * parseFloat(value) / 100, 0);
      if (format === 'xlsx') {
        await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
        await expect(page.locator('canvas[data-worker-zoom]').first()).toHaveAttribute('data-worker-zoom', String(parseFloat(value) / 100));
      }
    }
    await unchanged(page, format, info, original);
  });
}

test('docx: zoom step buttons walk presets and stop at both limits', async ({ page }, info) => {
  await open(page, 'docx');
  const original = await save(page, 'docx', info, 'original');
  const control = zoomControl(page, 'docx');
  const increase = page.getByRole('button', { name: 'Zoom in', exact: true });
  const decrease = page.getByRole('button', { name: 'Zoom out', exact: true });
  await preset(page, 'docx', '50%');
  await expect(decrease).toBeDisabled();
  for (const value of ['75%', '90%', '100%', '125%', '150%', '200%']) {
    await increase.click();
    await expect(control).toContainText(value);
  }
  await expect(increase).toBeDisabled();
  for (const value of ['150%', '125%', '100%', '90%', '75%', '50%']) {
    await decrease.click();
    await expect(control).toContainText(value);
  }
  await expect(decrease).toBeDisabled();
  await unchanged(page, 'docx', info, original);
});

test('docx: zoom preserves a selected word and its replacement can undo and redo', async ({ page }, info) => {
  await open(page, 'docx');
  await docxInput(page).press('ControlOrMeta+Home');
  for (let index = 0; index < 'Navigation'.length; index++) await docxInput(page).press('Shift+ArrowRight');
  await preset(page, 'docx', '150%');
  await docxInput(page).focus();
  await inside(docxInput(page), docxScroller(page));
  await page.keyboard.insertText(marker);
  await expect(page.locator('.canvas-page-mirror').first()).toContainText(`${marker} page 01`);
  const edited = await save(page, 'docx', info, 'edited');
  expect((await paragraphs(page, edited)).join('\n')).toContain(marker);
  await docxInput(page).press('ControlOrMeta+z');
  await expect(page.locator('.canvas-page-mirror').first()).toContainText('Navigation page 01');
  await expect(page.locator('.canvas-page-mirror').first()).not.toContainText(marker);
  await docxInput(page).press('ControlOrMeta+Shift+z');
  await expect(page.locator('.canvas-page-mirror').first()).toContainText(marker);
  const redone = await save(page, 'docx', info, 'redone');
  expect(redone.parts).toEqual(edited.parts);
  await open(page, 'docx', redone.file);
  await expect(page.locator('.canvas-page-mirror').first()).toContainText(marker);
  expect((await save(page, 'docx', info, 'reopened')).parts).toEqual(edited.parts);
});

test('docx: hit testing after zoom puts the caret at the clicked line', async ({ page }, info) => {
  await open(page, 'docx');
  await preset(page, 'docx', '150%');
  const paragraph = page.locator('.canvas-page-mirror').first().getByRole('paragraph').filter({ hasText: /^Page 01 line 3\./ });
  const line = paragraph.locator('.layout-run-text').first();
  await line.scrollIntoViewIfNeeded();
  await expect.poll(() => width(line)).toBeGreaterThan(0);
  await expect(line).toBeInViewport();
  const box = await settled(line);
  await page.mouse.click(box.x + Math.min(12, box.width / 2), box.y + box.height / 2);
  await docxInput(page).press('Home');
  await page.keyboard.insertText(marker);
  await expect(page.locator('.canvas-page-mirror').first()).toContainText(`${marker}Page 01 line 3.`);
  const edited = await save(page, 'docx', info, 'edited');
  expect(await paragraphs(page, edited)).toContain(`${marker}Page 01 line 3. The quick brown fox jumps over the lazy dog.`);
  await docxInput(page).press('ControlOrMeta+z');
  await expect(page.locator('.canvas-page-mirror').first()).not.toContainText(marker);
  await docxInput(page).press('ControlOrMeta+Shift+z');
  await expect(page.locator('.canvas-page-mirror').first()).toContainText(`${marker}Page 01 line 3.`);
  await docxInput(page).press('ControlOrMeta+z');
  await expect(page.locator('.canvas-page-mirror').first()).not.toContainText(marker);
  const undone = await save(page, 'docx', info, 'undone');
  expect((await paragraphs(page, undone)).join('\n')).not.toContain(marker);
  await open(page, 'docx', edited.file);
  await expect(page.locator('.canvas-page-mirror').first()).toContainText(`${marker}Page 01 line 3.`);
  expect((await save(page, 'docx', info, 'reopened')).parts).toEqual(edited.parts);
});

test('docx: Ctrl Home and End reveal and paint both document boundaries', async ({ page }, info) => {
  await open(page, 'docx');
  const original = await save(page, 'docx', info, 'original');
  await docxInput(page).press('ControlOrMeta+End');
  await expect(docxStatus(page)).toHaveText(`${pageCount} of ${pageCount}`);
  await inside(docxInput(page), docxScroller(page));
  await expect(page.locator('.canvas-page').filter({ hasText: 'Navigation page 16' })).toBeInViewport();
  await paintedDocx(page);
  await docxInput(page).press('ControlOrMeta+Home');
  await expect(docxStatus(page)).toHaveText(`1 of ${pageCount}`);
  await inside(docxInput(page), docxScroller(page));
  await paintedDocx(page);
  await unchanged(page, 'docx', info, original);
});

test('docx: PageDown and PageUp keep every landing painted and update the page indicator', async ({ page }, info) => {
  await open(page, 'docx');
  const original = await save(page, 'docx', info, 'original');
  const scroller = docxScroller(page);
  let previous = 0;
  for (let step = 0; step < 5; step++) {
    await docxInput(page).press('PageDown');
    await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(previous);
    previous = await scroller.evaluate((element) => element.scrollTop);
    await paintedDocx(page);
  }
  await expect.poll(async () => Number((await docxStatus(page).textContent())!.split(' ')[0])).toBeGreaterThan(2);
  await docxInput(page).press('PageUp');
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeLessThan(previous);
  await paintedDocx(page);
  await unchanged(page, 'docx', info, original);
});

test('docx: jumping the scroll position paints distant pages and returning to the top repaints', async ({ page }, info) => {
  await open(page, 'docx');
  const original = await save(page, 'docx', info, 'original');
  const scroller = docxScroller(page);
  await scroller.evaluate((element) => element.scrollTo({ top: (element.scrollHeight - element.clientHeight) * 0.8 }));
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop / (element.scrollHeight - element.clientHeight))).toBeGreaterThan(0.6);
  await expect.poll(async () => Number((await docxStatus(page).textContent())!.split(' ')[0])).toBeGreaterThan(8);
  await paintedDocx(page);
  await scroller.focus();
  await scroller.press('ControlOrMeta+Home');
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBe(0);
  await paintedDocx(page);
  await unchanged(page, 'docx', info, original);
});

test('xlsx: typed zoom preserves selection and hit testing targets the scaled cell', async ({ page }, info) => {
  await open(page, 'xlsx');
  const original = await save(page, 'xlsx', info, 'original');
  const scroll = page.getByTestId('xlsx-scroll');
  for (const key of ['ArrowRight', 'ArrowRight', 'ArrowDown', 'ArrowDown']) await scroll.press(key);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('C3');
  const selection = page.getByTestId('xlsx-selection');
  const initial = (await settled(selection)).width;
  await typedZoom(page, 'xlsx', '137%');
  await expect(page.locator('canvas[data-worker-zoom]').first()).toHaveAttribute('data-worker-zoom', '1.37');
  await expect.poll(() => width(selection)).toBeCloseTo(initial * 1.37, 0);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('C3');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Hit target 1');
  await inside(selection, scroll);
  const target = await settled(selection);
  const offset = await scroll.evaluate((element) => [element.scrollLeft, element.scrollTop]);
  await scroll.press('ControlOrMeta+Home');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  expect(await scroll.evaluate((element) => [element.scrollLeft, element.scrollTop])).toEqual(offset);
  await page.mouse.click(target.x + target.width / 2, target.y + target.height / 2);
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('C3');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Hit target 1');
  await unchanged(page, 'xlsx', info, original);
});

test('xlsx: zoom persists across sheet switches and viewport resize', async ({ page }, info) => {
  await open(page, 'xlsx');
  const original = await save(page, 'xlsx', info, 'original');
  await typedZoom(page, 'xlsx', '137%');
  for (const [sheet, value] of [['Formulas', 'Origin 2'], ['Data', 'Origin 3'], ['Dashboard', 'Origin 1']] as const) {
    const tab = page.getByRole('tab', { name: sheet, exact: true });
    await tab.click();
    await expect(tab).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(value);
    await expect(zoomControl(page, 'xlsx')).toHaveValue('137%');
    await expect(page.locator('canvas[data-worker-zoom]').first()).toHaveAttribute('data-worker-zoom', '1.37');
  }
  await page.setViewportSize({ width: 1100, height: 760 });
  await inside(page.getByTestId('xlsx-selection'), page.getByTestId('xlsx-scroll'));
  await expect(zoomControl(page, 'xlsx')).toHaveValue('137%');
  await unchanged(page, 'xlsx', info, original);
});

test('xlsx: keyboard navigation reaches row 10000 and column ZZ and returns to A1', async ({ page }, info) => {
  await open(page, 'xlsx');
  const original = await save(page, 'xlsx', info, 'original');
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.press('ControlOrMeta+End');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('ZZ10000');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Far corner 1');
  await expect(page.getByRole('gridcell').filter({ hasText: 'Far corner 1' })).toHaveCount(1);
  await inside(page.getByTestId('xlsx-selection'), scroll);
  await preset(page, 'xlsx', '150%');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('ZZ10000');
  await inside(page.getByTestId('xlsx-selection'), scroll);
  await scroll.press('ControlOrMeta+Home');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Origin 1');
  await inside(page.getByTestId('xlsx-selection'), scroll);
  for (const [key, address, value] of [
    ['ArrowDown', 'A10000', 'Far row 1'],
    ['ArrowRight', 'ZZ10000', 'Far corner 1'],
    ['ArrowUp', 'ZZ1', 'Far column 1'],
    ['ArrowLeft', 'A1', 'Origin 1'],
  ] as const) {
    await scroll.press(`ControlOrMeta+${key}`);
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue(address);
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(value);
    await inside(page.getByTestId('xlsx-selection'), scroll);
  }
  await unchanged(page, 'xlsx', info, original);
});

test('pptx: typed zoom preserves the selected shape and hit testing follows its scaled bounds', async ({ page }, info) => {
  await open(page, 'pptx');
  const original = await save(page, 'pptx', info, 'original');
  const canvas = page.getByTestId('pptx-slide-canvas');
  await preset(page, 'pptx', '100%');
  await expect.poll(() => width(page.getByTestId('pptx-slide-canvas'))).toBeCloseTo(1280, 0);
  const fullWidth = (await settled(canvas)).width;
  await preset(page, 'pptx', '50%');
  await expect.poll(() => width(canvas)).toBeCloseTo(fullWidth / 2, 0);
  let bounds = await settled(canvas);
  await page.mouse.click(bounds.x + bounds.width * 0.35, bounds.y + bounds.height * 0.45 - 2);
  const selection = page.getByTestId('pptx-shape-selection');
  await expect(selection).toBeVisible();
  const initial = (await settled(selection)).width;
  await typedZoom(page, 'pptx', '83%');
  await expect.poll(() => width(selection)).toBeCloseTo(initial * 0.83 / 0.5, 0);
  const stage = stageOf(canvas);
  await inside(selection, stage);
  await stage.press('PageDown');
  await expect(page.getByRole('complementary', { name: 'Slides', exact: true }).locator('button').nth(1)).toHaveAttribute('aria-current', 'page');
  await stage.press('PageUp');
  await expect(page.getByRole('complementary', { name: 'Slides', exact: true }).locator('button').first()).toHaveAttribute('aria-current', 'page');
  await expect(selection).toHaveCount(0);
  bounds = await settled(canvas);
  await page.mouse.click(bounds.x + bounds.width * 0.35, bounds.y + bounds.height * 0.45 - 2);
  await expect(selection).toBeVisible();
  await expect.poll(async () => {
    const shape = await selection.boundingBox();
    return shape ? Math.abs(shape.width - bounds.width * 0.3) + Math.abs(shape.height - bounds.height * 0.25) +
      Math.abs(shape.x - (bounds.x + bounds.width * 0.2)) + Math.abs(shape.y - (bounds.y + bounds.height * 0.2)) : Infinity;
  }).toBeLessThan(5);
  await unchanged(page, 'pptx', info, original);
});

test('pptx: scrolling a forty-slide strip selects the last slide and retains zoom', async ({ page }, info) => {
  await open(page, 'pptx', undefined, 40);
  const original = await save(page, 'pptx', info, 'original');
  await preset(page, 'pptx', '100%');
  await expect.poll(() => width(page.getByTestId('pptx-slide-canvas'))).toBeCloseTo(1280, 0);
  const fullWidth = (await settled(page.getByTestId('pptx-slide-canvas'))).width;
  await typedZoom(page, 'pptx', '75%');
  await expect.poll(() => width(page.getByTestId('pptx-slide-canvas'))).toBeCloseTo(fullWidth * 0.75, 0);
  const strip = page.getByRole('complementary', { name: 'Slides', exact: true });
  const buttons = strip.locator('button');
  const canvas = page.getByTestId('pptx-slide-canvas');
  const first = await canvas.screenshot();
  await strip.hover();
  await page.mouse.wheel(0, 100_000);
  await expect(buttons.last()).toBeInViewport();
  await buttons.last().click();
  await expect(buttons.last()).toHaveAttribute('aria-current', 'page');
  await expect(zoomControl(page, 'pptx')).toHaveValue('75%');
  await expect.poll(async () => (await canvas.screenshot()).equals(first)).toBe(false);
  await stageOf(canvas).press('Home');
  await expect(buttons.first()).toHaveAttribute('aria-current', 'page');
  await expect(zoomControl(page, 'pptx')).toHaveValue('75%');
  await stageOf(canvas).press('End');
  await expect(buttons.last()).toHaveAttribute('aria-current', 'page');
  await expect(zoomControl(page, 'pptx')).toHaveValue('75%');
  await unchanged(page, 'pptx', info, original);
});

test('pptx: clicking the slide then PageDown and PageUp move the active slide', async ({ page }) => {
  await open(page, 'pptx');
  const canvas = page.getByTestId('pptx-slide-canvas');
  const buttons = page.getByRole('complementary', { name: 'Slides', exact: true }).locator('button');
  const box = await settled(canvas);
  await page.mouse.click(box.x + box.width * 0.9, box.y + box.height * 0.9);
  await page.keyboard.press('PageDown');
  await expect(buttons.nth(1)).toHaveAttribute('aria-current', 'page');
  await page.keyboard.press('PageUp');
  await expect(buttons.first()).toHaveAttribute('aria-current', 'page');
});

test('pptx: Fit responds to window resize and slide navigation without changing saved OOXML', async ({ page }, info) => {
  await open(page, 'pptx');
  const original = await save(page, 'pptx', info, 'original');
  const canvas = page.getByTestId('pptx-slide-canvas');
  await preset(page, 'pptx', '100%');
  await preset(page, 'pptx', 'Fit');
  await fitted(canvas);
  const initial = (await settled(canvas)).width;
  await page.setViewportSize({ width: 1100, height: 760 });
  await expect.poll(() => width(canvas)).toBeLessThan(initial);
  await fitted(canvas);
  await expect(zoomControl(page, 'pptx')).toHaveValue('Fit');
  await inside(canvas, stageOf(canvas));
  await stageOf(canvas).press('PageDown');
  const buttons = page.getByRole('complementary', { name: 'Slides', exact: true }).locator('button');
  await expect(buttons.nth(1)).toHaveAttribute('aria-current', 'page');
  await fitted(canvas);
  await expect(zoomControl(page, 'pptx')).toHaveValue('Fit');
  await inside(canvas, stageOf(canvas));
  await stageOf(canvas).press('PageUp');
  await expect(buttons.first()).toHaveAttribute('aria-current', 'page');
  await page.setViewportSize({ width: 1440, height: 1000 });
  await expect.poll(() => width(canvas)).toBeCloseTo(initial, 0);
  await unchanged(page, 'pptx', info, original);
});
