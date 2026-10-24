import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

type Format = 'docx' | 'xlsx' | 'pptx';
type BrowserProbe = {
  workers: { worker: Worker; url: string; active: boolean }[];
  printed: number;
  popups: Window[];
  painted(): string;
};

declare global {
  interface Window {
    __asyncRobustness: BrowserProbe;
  }
}

const formats = ['docx', 'xlsx', 'pptx'] as const;
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relationships = (body: string) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

async function fixture(format: Format) {
  const zip = new JSZip();
  const main = format === 'docx' ? 'word/document.xml' : format === 'xlsx' ? 'xl/workbook.xml' : 'ppt/presentation.xml';
  const overrides: [string, string][] = [];
  if (format === 'docx') {
    overrides.push([main, 'wordprocessingml.document.main']);
    zip.file(main, '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Seed</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  } else if (format === 'xlsx') {
    const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    overrides.push([main, 'spreadsheetml.sheet.main'], ['xl/worksheets/sheet1.xml', 'spreadsheetml.worksheet']);
    zip.file(main, `<workbook xmlns="${ns}" xmlns:r="${office}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', relationships(`<Relationship Id="rId1" Type="${office}/worksheet" Target="worksheets/sheet1.xml"/>`));
    zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${ns}"><dimension ref="A1:B2"/><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="2" width="14" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Seed</t></is></c><c r="B1" t="inlineStr"><is><t>Second</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row></sheetData></worksheet>`);
  } else {
    const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    overrides.push([main, 'presentationml.presentation.main'], ['ppt/slides/slide1.xml', 'presentationml.slide']);
    zip.file(main, `<p:presentation xmlns:p="${p}" xmlns:r="${office}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
    zip.file('ppt/_rels/presentation.xml.rels', relationships(`<Relationship Id="rId1" Type="${office}/slide" Target="slides/slide1.xml"/>`));
    zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Editable text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="6400800" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Seed</a:t></a:r><a:endParaRPr sz="2400"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  }
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + overrides.map(([path, type]) => `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`).join('') + '</Types>');
  zip.file('_rels/.rels', relationships(`<Relationship Id="rId1" Type="${office}/officeDocument" Target="${main}"/>`));
  return { name: `robustness.${format}`, mimeType: 'application/octet-stream', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function instrument(page: Page) {
  await page.addInitScript(() => {
    const texts = new WeakMap<HTMLCanvasElement, string[]>();
    const probe: BrowserProbe = {
      workers: [], printed: 0, popups: [],
      painted() {
        const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="pptx-slide-canvas"]');
        return canvas ? (texts.get(canvas) ?? []).join('') : '';
      },
    };
    window.__asyncRobustness = probe;
    window.Worker = new Proxy(window.Worker, {
      construct(target, args) {
        const worker = Reflect.construct(target, args) as Worker;
        const record = { worker, url: String(args[0]), active: true };
        const terminate = worker.terminate.bind(worker);
        worker.terminate = () => { record.active = false; terminate(); };
        probe.workers.push(record);
        return worker;
      },
    });
    const fillText = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (...args: Parameters<typeof fillText>) {
      const values = texts.get(this.canvas) ?? [];
      values.push(args[0]);
      texts.set(this.canvas, values.slice(-256));
      fillText.apply(this, args);
    };
    const clearRect = CanvasRenderingContext2D.prototype.clearRect;
    CanvasRenderingContext2D.prototype.clearRect = function (...args: Parameters<typeof clearRect>) {
      if (args[0] === 0 && args[1] === 0) texts.set(this.canvas, []);
      clearRect.apply(this, args);
    };
    const drawImage = CanvasRenderingContext2D.prototype.drawImage;
    CanvasRenderingContext2D.prototype.drawImage = function (source: CanvasImageSource, ...args: number[]) {
      const copied = texts.get(source as HTMLCanvasElement);
      if (copied) texts.set(this.canvas, [...(texts.get(this.canvas) ?? []), ...copied].slice(-256));
      Reflect.apply(drawImage, this, [source, ...args]);
    };
    window.print = () => { probe.printed += 1; };
    const open = window.open.bind(window);
    window.open = (...args: Parameters<typeof window.open>) => {
      const popup = open(...args);
      if (popup) {
        popup.print = () => { probe.printed += 1; };
        popup.close = () => {};
        probe.popups.push(popup);
      }
      return popup;
    };
  });
}

const docxInput = (page: Page) => page.locator('.paged-editor__input');
const slideInput = (page: Page) => page.getByRole('textbox', { name: 'Slide text', exact: true });
const canvas = (page: Page, format: Format) => format === 'pptx'
  ? page.getByTestId('pptx-slide-canvas') : page.locator(format === 'docx' ? '.canvas-page canvas' : '[data-testid="xlsx-scroll"] canvas').first();

async function ready(page: Page, format: Format) {
  const layer = page.locator('.editor-layer');
  await expect(layer).toHaveCount(1);
  await expect(layer).toHaveAttribute('aria-hidden', 'false');
  await expect(canvas(page, format)).toBeVisible();
  await expect(page.locator('.notice')).toHaveCount(0);
  if (format === 'docx') await expect(page.locator('.canvas-page-mirror').first()).toContainText(/\S/);
  if (format === 'xlsx') await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(/.+/);
  if (format === 'pptx') {
    await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
    await expect.poll(() => page.evaluate(() => window.__asyncRobustness.painted())).toMatch(/\S/);
  }
}

async function open(page: Page, format: Format, file?: string) {
  await page.goto(`/?format=${format}`);
  await page.locator('input[type=file]').last().setInputFiles(file ?? await fixture(format));
  await ready(page, format);
}

async function edit(page: Page, format: Format, value: string) {
  if (format === 'docx') {
    await canvas(page, format).click({ position: { x: 110, y: 106 } });
    await docxInput(page).press('ControlOrMeta+Home');
    await page.keyboard.insertText(value);
  } else if (format === 'xlsx') {
    const scroll = page.getByTestId('xlsx-scroll');
    await scroll.press('ControlOrMeta+Home');
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await scroll.press('F2');
    const input = page.getByTestId('xlsx-cell-editor');
    await expect(input).toHaveValue('Seed');
    await input.press('ControlOrMeta+A');
    await page.keyboard.insertText(value);
    await input.press('Enter');
  } else {
    const box = await canvas(page, format).boundingBox();
    expect(box).not.toBeNull();
    await page.mouse.click(box!.x + box!.width * 0.12, box!.y + box!.height * 0.22);
    await expect(slideInput(page)).toBeFocused();
    await slideInput(page).press('Home');
    await page.keyboard.insertText(value);
  }
  await visible(page, format, value);
}

async function visible(page: Page, format: Format, value: string) {
  if (format === 'docx') await expect(page.locator('.canvas-page-mirror').first()).toContainText(value);
  else if (format === 'xlsx') await expect(page.getByRole('gridcell').filter({ hasText: value })).toHaveCount(1);
  else await expect.poll(() => page.evaluate(() => window.__asyncRobustness.painted())).toContain(value);
}

async function save(page: Page, format: Format, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${label}.${format}`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const parts = Object.values(zip.files).filter((part) => !part.dir && part.name.endsWith('.xml'))
    .sort((a, b) => a.name.localeCompare(b.name));
  const xml = await Promise.all(parts.map((part) => part.async('string')));
  const text = await page.evaluate((parts) => parts.map((xml) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved XML');
    return Array.from(document.getElementsByTagNameNS('*', 't')).map((node) => node.textContent ?? '').join('');
  }).join('\n'), xml);
  await expect(page.locator('.document-name')).not.toContainText('Edited');
  return { file, text };
}

async function history(page: Page, format: Format, info: TestInfo, value: string, original: string) {
  if (format === 'docx') await docxInput(page).press('ControlOrMeta+z');
  else await page.getByTestId(`${format}-undo`).click();
  if (format === 'docx') await expect(page.locator('.canvas-page-mirror').first()).not.toContainText(value);
  else if (format === 'xlsx') await expect(page.getByRole('gridcell').filter({ hasText: value })).toHaveCount(0);
  else await expect.poll(() => page.evaluate(() => window.__asyncRobustness.painted())).not.toContain(value);
  const undone = await save(page, format, info, 'undone');
  expect(undone.text).toBe(original);
  if (format === 'docx') await docxInput(page).press('ControlOrMeta+Shift+z');
  else await page.getByTestId(`${format}-redo`).click();
  await visible(page, format, value);
  const redone = await save(page, format, info, 'redone');
  expect(redone.text).toContain(value);
  await open(page, format, redone.file);
  expect((await save(page, format, info, 'reopened')).text).toBe(redone.text);
}

test.beforeEach(async ({ page }) => { await instrument(page); });

for (const format of ['docx', 'pptx'] as const) {
  test(`${format}: worker crash recovery preserves edits and accepts new typing`, async ({ page }, info) => {
    await open(page, format);
    await edit(page, format, 'BeforeCrash');
    const workerName = format === 'docx' ? 'docxWorker' : `${format}SessionWorker`;
    const count = await page.evaluate(() => window.__asyncRobustness.workers.length);
    await page.evaluate((name) => {
      const record = [...window.__asyncRobustness.workers].reverse().find((entry) => entry.active && entry.url.includes(name));
      if (!record) throw new Error(`No active ${name} worker`);
      record.worker.dispatchEvent(new ErrorEvent('error', { message: 'Browser worker failure', error: new Error('Browser worker failure') }));
      record.worker.terminate();
    }, workerName);
    if (format === 'pptx') {
      const downloading = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Download recovered presentation', exact: true }).click();
      const recovered = info.outputPath('recovered.pptx');
      await (await downloading).saveAs(recovered);
      const zip = await JSZip.loadAsync(await readFile(recovered));
      expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain('BeforeCrash');
      const confirming = page.waitForEvent('dialog').then(async (dialog) => {
        expect(dialog.type()).toBe('confirm');
        await dialog.accept();
      });
      await page.locator('input[type=file]').last().setInputFiles(recovered);
      await confirming;
      await ready(page, format);
    }
    if (format === 'docx') {
      await expect.poll(() => page.evaluate(() => window.__asyncRobustness.workers.length)).toBeGreaterThan(count);
    }
    await expect(page.locator('.notice, [data-testid$="-error"]')).toHaveCount(0);
    await visible(page, format, 'BeforeCrash');
    if (format === 'docx') {
      await docxInput(page).press('ControlOrMeta+z');
      await expect(page.locator('.canvas-page-mirror').first()).toHaveText('Seed');
      await docxInput(page).press('ControlOrMeta+Shift+z');
      await visible(page, format, 'BeforeCrash');
    }
    await edit(page, format, 'AfterCrash');
    await visible(page, format, 'AfterCrash');
    const saved = await save(page, format, info, 'after-crash');
    expect(saved.text).toContain('BeforeCrash');
    expect(saved.text).toContain('AfterCrash');
    await history(page, format, info, 'AfterCrash', saved.text.replace('AfterCrash', ''));
  });
}

for (const format of formats) {
  test(`${format}: closing an edited file prompts and saving permits a clean reload`, async ({ page }, info) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    page.on('console', (message) => { if (message.type() === 'error') errors.push(message.text()); });
    await open(page, format);
    await edit(page, format, 'UnsavedProbe');
    await expect(page.locator('.document-name')).toContainText('Edited');
    const confirming = page.waitForEvent('dialog').then(async (dialog) => {
      expect(dialog.type()).toBe('confirm');
      expect(dialog.message()).toBe('Discard unsaved changes?');
      await dialog.dismiss();
    });
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await confirming;
    await visible(page, format, 'UnsavedProbe');
    const unloading = page.waitForEvent('dialog');
    const cancelledReload = page.reload().catch(() => null);
    const beforeUnload = await unloading;
    expect(beforeUnload.type()).toBe('beforeunload');
    await beforeUnload.dismiss();
    await cancelledReload;
    await visible(page, format, 'UnsavedProbe');
    const saved = await save(page, format, info, 'before-reload');
    expect(saved.text).toContain('UnsavedProbe');
    const dialogs: string[] = [];
    page.on('dialog', async (dialog) => { dialogs.push(dialog.type()); await dialog.dismiss(); });
    await page.reload();
    await expect(page.locator('.editor-stage canvas')).toHaveCount(0);
    await page.locator('input[type=file]').last().setInputFiles(saved.file);
    await ready(page, format);
    expect((await save(page, format, info, 'after-reload')).text).toBe(saved.text);
    expect(dialogs).toEqual([]);
    expect(errors).toEqual([]);
  });

  test(`${format}: print requests the browser dialog without errors and preserves saved text`, async ({ page }, info) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    page.on('console', (message) => { if (message.type() === 'error') errors.push(message.text()); });
    await open(page, format);
    await edit(page, format, 'PrintProbe');
    const before = await save(page, format, info, 'before-print');
    await page.keyboard.press('ControlOrMeta+p');
    await expect.poll(() => page.evaluate(() => window.__asyncRobustness.printed)).toBe(1);
    if (format === 'docx') {
      await expect.poll(() => page.evaluate(() => {
        const images = window.__asyncRobustness.popups[0]?.document.images;
        return !!images?.length && [...images].every((image) => image.complete && image.naturalWidth > 0);
      })).toBe(true);
    }
    expect((await save(page, format, info, 'after-print')).text).toBe(before.text);
    await open(page, format, before.file);
    expect((await save(page, format, info, 'print-reopened')).text).toBe(before.text);
    await expect(page.locator('.notice, [data-testid$="-error"]')).toHaveCount(0);
    expect(errors).toEqual([]);
  });
}

test.describe('touch input', () => {
  test.use({ hasTouch: true });

  for (const format of formats) {
    test(`${format}: touch selection accepts keyboard edits through undo redo and reopen`, async ({ page }, info) => {
      await open(page, format);
      const original = await save(page, format, info, 'touch-original');
      const box = await canvas(page, format).boundingBox();
      expect(box).not.toBeNull();
      if (format === 'docx') {
        await page.touchscreen.tap(box!.x + box!.width * 110 / 816, box!.y + box!.width * 106 / 816);
        await expect(docxInput(page)).toBeFocused();
        await page.keyboard.insertText('TouchProbe');
      } else if (format === 'xlsx') {
        const selection = await page.getByTestId('xlsx-selection').boundingBox();
        expect(selection).not.toBeNull();
        await page.touchscreen.tap(selection!.x + selection!.width * 1.5, selection!.y + selection!.height / 2);
        await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B1');
        await page.keyboard.press('F2');
        await expect(page.getByTestId('xlsx-cell-editor')).toHaveValue('Second');
        await page.keyboard.press('ControlOrMeta+A');
        await page.keyboard.insertText('TouchProbe');
        await page.keyboard.press('Enter');
        await expect(page.getByTestId('xlsx-name-box')).toHaveValue('B2');
      } else {
        await page.touchscreen.tap(box!.x + box!.width * 0.1 + 2, box!.y + box!.height * 0.35);
        await expect(page.getByTestId('pptx-shape-selection')).toBeVisible();
        await page.touchscreen.tap(box!.x + box!.width * 0.12, box!.y + box!.height * 0.22);
        await expect(slideInput(page)).toBeFocused();
        await page.keyboard.insertText('TouchProbe');
        await expect(page.getByTestId('pptx-peerless-caret')).toBeVisible();
      }
      await visible(page, format, 'TouchProbe');
      expect((await save(page, format, info, 'touch-edited')).text).toContain('TouchProbe');
      await history(page, format, info, 'TouchProbe', original.text);
    });
  }
});

for (const format of ['xlsx', 'pptx'] as const) {
  test(`${format}: Export PNG downloads a decodable image at the expected size`, async ({ page }, info) => {
    await open(page, format);
    await edit(page, format, 'PngProbe');
    const before = await save(page, format, info, 'before-png');
    const size = format === 'pptx' ? { width: 960, height: 540 }
      : await page.getByTestId('xlsx-scroll').evaluate((element) => ({ width: element.clientWidth, height: element.clientHeight }));
    const downloading = page.waitForEvent('download');
    await page.getByTestId(`${format}-export-png`).click();
    const download = await downloading;
    expect(download.suggestedFilename()).toMatch(/\.png$/i);
    const file = info.outputPath(`${format}-export.png`);
    await download.saveAs(file);
    const bytes = await readFile(file);
    expect([...bytes.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(bytes.subarray(12, 16).toString('ascii')).toBe('IHDR');
    expect({ width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) }).toEqual(size);
    const decoded = await page.evaluate(async (base64) => {
      const blob = await (await fetch(`data:image/png;base64,${base64}`)).blob();
      const image = await createImageBitmap(blob);
      const result = { width: image.width, height: image.height };
      image.close();
      return result;
    }, bytes.toString('base64'));
    expect(decoded).toEqual(size);
    expect((await save(page, format, info, 'after-png')).text).toBe(before.text);
    await open(page, format, before.file);
    expect((await save(page, format, info, 'png-reopened')).text).toBe(before.text);
  });
}
