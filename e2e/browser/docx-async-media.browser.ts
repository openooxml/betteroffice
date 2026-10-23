import { test, expect, type Locator, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { posix } from 'node:path';
import JSZip from 'jszip';

const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const imageName = 'green-picture.png';
const emu = 9525;
const input = (page: Page) => page.locator('.paged-editor__peerless-input');
const pictures = (page: Page) => page.locator('.canvas-page-mirror .layout-run-image');
const mirror = (page: Page) => page.locator('.canvas-page-mirror').first();
const handles = (page: Page) => page.locator('.canvas-image-selection-overlay [data-handle]');

test.describe.configure({ timeout: 30_000 });

async function raster(page: Page, extension: 'png' | 'jpeg' = 'png') {
  const data = await page.evaluate(extension => {
    const canvas = document.createElement('canvas');
    canvas.width = 120;
    canvas.height = 60;
    const context = canvas.getContext('2d')!;
    context.fillStyle = '#14a078';
    context.fillRect(0, 0, canvas.width, canvas.height);
    return canvas.toDataURL(`image/${extension}`, 0.95).split(',')[1]!;
  }, extension);
  return { name: extension === 'png' ? imageName : 'green-picture.jpeg', mimeType: `image/${extension}`, buffer: Buffer.from(data, 'base64') };
}

function picture(floating: boolean, behind = false) {
  const content = `<wp:extent cx="1143000" cy="571500"/><wp:docPr id="1" name="Picture 1" descr="${imageName}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="571500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  return '<w:r><w:drawing>' + (floating
    ? `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="${behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>1905000</wp:posOffset></wp:positionV>${content.replace('<wp:docPr', '<wp:wrapNone/><wp:docPr')}</wp:anchor>`
    : `<wp:inline distT="0" distB="0" distL="0" distR="0">${content}</wp:inline>`) + '</w:drawing></w:r>';
}

function textBox(shape: boolean) {
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2286000" cy="914400"/><wp:docPr id="2" name="Text box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="${shape ? '0' : '1'}"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="914400"/></a:xfrm><a:prstGeom prst="${shape ? 'roundRect' : 'rect'}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EEF3FF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="2F5496"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p w14:paraId="00000003"><w:r><w:t>Box seed</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

async function fixture(page: Page, content: 'text' | 'inline' | 'floating' | 'behind' | 'textbox' | 'shape' = 'inline') {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const run = content === 'text' ? '<w:r><w:t>BeforeAfter</w:t></w:r>'
    : content === 'textbox' || content === 'shape' ? textBox(content === 'shape')
    : picture(content !== 'inline', content === 'behind') + '<w:r><w:t>Tail</w:t></w:r>';
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body><w:p w14:paraId="00000001">${run}</w:p><w:p w14:paraId="00000002"><w:r><w:t>Destination</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  if (content === 'inline' || content === 'floating' || content === 'behind') {
    zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/></Relationships>');
    zip.file('word/media/picture.png', (await raster(page)).buffer);
  }
  return { name: 'media.docx', mimeType, buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function open(page: Page, file?: string | Awaited<ReturnType<typeof fixture>>, content: Parameters<typeof fixture>[1] = 'inline') {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles(file ?? await fixture(page, content));
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible();
  await expect(input(page)).toBeEditable();
  await expect(page.getByRole('button', { name: 'Save', exact: true }).first()).toBeEnabled();
  await expect(page.locator('.notice')).toHaveCount(0);
  await input(page).focus();
  await input(page).press('ControlOrMeta+Home');
}

async function save(page: Page, info: TestInfo, name: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  const state = await page.evaluate(xml => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved document XML');
    const elements = (node: Element | Document, name: string) => Array.from(node.getElementsByTagNameNS('*', name));
    const body = elements(document, 'body')[0]!;
    const text = (node: Element): string => node.localName === 'drawing' ? '\uFFFC'
      : node.localName === 't' ? node.textContent ?? '' : Array.from(node.children).map(text).join('');
    return {
      paragraphs: Array.from(body.children).filter(node => node.localName === 'p').map(text),
      boxes: elements(body, 'txbxContent').map(box => elements(box, 't').map(node => node.textContent).join('')),
      drawings: elements(body, 'drawing').map(drawing => {
        const container = drawing.firstElementChild!;
        const extent = elements(container, 'extent')[0]!;
        const wrap = Array.from(container.children).find(node => node.localName.startsWith('wrap'));
        return {
          kind: container.localName,
          width: Number(extent.getAttribute('cx')),
          height: Number(extent.getAttribute('cy')),
          alt: elements(container, 'docPr')[0]?.getAttribute('descr') ?? '',
          embed: elements(container, 'blip')[0]?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') ?? '',
          wrap: wrap?.localName ?? '',
          side: wrap?.getAttribute('wrapText') ?? '',
          behind: container.getAttribute('behindDoc') ?? '',
          offsets: elements(container, 'posOffset').map(node => Number(node.textContent)),
        };
      }),
    };
  }, xml);
  return { file, zip, state };
}

type Saved = Awaited<ReturnType<typeof save>>;

async function expectMedia(page: Page, saved: Saved, expected: Buffer) {
  const relationships = await saved.zip.file('word/_rels/document.xml.rels')!.async('string');
  const contentTypes = await saved.zip.file('[Content_Types].xml')!.async('string');
  const links = await page.evaluate(({ relationships, contentTypes }) => {
    const parse = (xml: string) => {
      const document = new DOMParser().parseFromString(xml, 'application/xml');
      if (document.querySelector('parsererror')) throw new Error('Invalid saved package XML');
      return document;
    };
    const types = parse(contentTypes);
    return Array.from(parse(relationships).getElementsByTagNameNS('*', 'Relationship')).map(node => {
      const target = node.getAttribute('Target')!;
      const extension = target.split('.').pop();
      const override = Array.from(types.getElementsByTagNameNS('*', 'Override')).find(type => type.getAttribute('PartName') === `/word/${target}`);
      const fallback = Array.from(types.getElementsByTagNameNS('*', 'Default')).find(type => type.getAttribute('Extension') === extension);
      return { id: node.getAttribute('Id'), target, type: node.getAttribute('Type'), external: node.getAttribute('TargetMode'), mimeType: (override ?? fallback)?.getAttribute('ContentType') };
    });
  }, { relationships, contentTypes });
  for (const drawing of saved.state.drawings.filter(drawing => drawing.embed)) {
    const link = links.find(link => link.id === drawing.embed);
    expect(link).toBeDefined();
    expect(link!.type).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
    expect(link!.external).not.toBe('External');
    expect(link!.mimeType).toBe(expected[0] === 0xff ? 'image/jpeg' : 'image/png');
    const part = saved.zip.file(posix.resolve('/word', link!.target).slice(1));
    expect(part).not.toBeNull();
    expect(await part!.async('nodebuffer')).toEqual(expected);
  }
}

async function expectPainted(page: Page, image: Locator = pictures(page).first()) {
  await expect(image).toBeAttached();
  await expect.poll(async () => {
    const box = await image.boundingBox();
    if (!box) return false;
    const screenshot = await page.screenshot({ clip: { x: Math.floor(box.x + box.width / 2), y: Math.floor(box.y + box.height / 2), width: 1, height: 1 } });
    return page.evaluate(async data => {
      const image = new Image();
      image.src = `data:image/png;base64,${data}`;
      await image.decode();
      const canvas = document.createElement('canvas');
      canvas.width = canvas.height = 1;
      const context = canvas.getContext('2d')!;
      context.drawImage(image, 0, 0);
      const pixel = context.getImageData(0, 0, 1, 1).data;
      const original = [20, 160, 120];
      const selectionTint = [59, 130, 246];
      const selected = original.map((channel, index) => channel * 0.75 + selectionTint[index]! * 0.25);
      return [original, selected].some(color => color.every((channel, index) => Math.abs(pixel[index]! - channel) < 12));
    }, screenshot.toString('base64'));
  }).toBe(true);
}

async function imageBounds(page: Page) {
  let box = await pictures(page).first().boundingBox();
  await expect.poll(async () => {
    box = await pictures(page).first().boundingBox();
    return box;
  }).not.toBeNull();
  return box!;
}

async function selectByKeyboard(page: Page) {
  await input(page).focus();
  await input(page).press('ControlOrMeta+Home');
  await input(page).press('Shift+ArrowRight');
}

async function history(page: Page, info: TestInfo, before: Saved, after: Saved, visibleBefore: () => Promise<void>, visibleAfter: () => Promise<void>) {
  await input(page).focus();
  await input(page).press('ControlOrMeta+z');
  await visibleBefore();
  expect((await save(page, info, 'undone')).state).toEqual(before.state);
  await input(page).focus();
  await input(page).press('ControlOrMeta+Shift+z');
  await visibleAfter();
  const redone = await save(page, info, 'redone');
  expect(redone.state).toEqual(after.state);
  return redone;
}

async function reopen(page: Page, info: TestInfo, saved: Saved) {
  await open(page, saved.file);
  await expect(pictures(page)).toHaveCount(saved.state.drawings.filter(drawing => drawing.embed).length);
  const reopened = await save(page, info, 'reopened');
  expect(reopened.state).toEqual(saved.state);
  return reopened;
}

async function properties(page: Page) {
  const button = page.getByRole('button', { name: 'Image properties', exact: true });
  await expect(button).toBeAttached();
  await expect(button).not.toHaveAttribute('aria-disabled', 'true');
  if (await button.isVisible()) await button.click();
  else {
    await page.getByRole('button', { name: 'More actions', exact: true }).click();
    await page.getByRole('menuitem', { name: 'Image properties', exact: true }).click();
  }
  const dialog = page.getByRole('dialog', { name: 'Image Properties', exact: true });
  await expect(dialog).toBeVisible();
  return dialog;
}

async function wrap(page: Page, name: string) {
  const box = await imageBounds(page);
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' });
  await page.locator('.docx-text-context-menu').getByRole('menuitem', { name, exact: true }).click();
  await expect(page.locator('.docx-text-context-menu')).toBeHidden();
}

for (const extension of ['png', 'jpeg'] as const) {
  test(`Insert menu places a ${extension.toUpperCase()} at the caret and preserves its media after reopen`, async ({ page }, info) => {
    await open(page, undefined, 'text');
    const before = await save(page, info, 'before-insert');
    await input(page).focus();
    await input(page).press('ControlOrMeta+Home');
    for (let index = 0; index < 6; index++) await input(page).press('ArrowRight');
    const image = await raster(page, extension);
    await page.getByRole('menubar').getByRole('menuitem', { name: 'Insert', exact: true }).click();
    const choosing = page.waitForEvent('filechooser');
    await page.getByRole('menu', { name: 'Insert', exact: true }).getByRole('menuitem', { name: 'Image', exact: true }).click();
    await (await choosing).setFiles(image);
    await expect(pictures(page)).toHaveCount(1);
    await expect(pictures(page).first()).toHaveAttribute('aria-label', image.name);
    await expectPainted(page);
    const bounds = await imageBounds(page);
    // The worker mirror paints one text span per shaped cluster, so locate the paragraph's first and last glyphs.
    const runs = mirror(page).locator('.layout-paragraph').filter({ hasText: /^Before.*After$/ }).locator('.layout-run-text');
    const prefix = runs.first();
    const suffix = runs.last();
    await expect(prefix).toBeAttached();
    await expect(suffix).toBeAttached();
    expect((await prefix.boundingBox())!.x).toBeLessThan(bounds.x);
    expect((await suffix.boundingBox())!.x).toBeGreaterThanOrEqual(bounds.x + bounds.width - 1);
    const inserted = await save(page, info, 'inserted');
    expect(inserted.state.paragraphs).toEqual(['Before\uFFFCAfter', 'Destination']);
    expect(inserted.state.drawings).toHaveLength(1);
    expect(inserted.state.drawings[0]).toMatchObject({ kind: 'inline', width: 120 * emu, height: 60 * emu, alt: image.name });
    await expectMedia(page, inserted, image.buffer);
    const redone = await history(page, info, before, inserted,
      () => expect(pictures(page)).toHaveCount(0), () => expectPainted(page));
    const reopened = await reopen(page, info, redone);
    await expectPainted(page);
    await expectMedia(page, reopened, image.buffer);
  });
}

test('clicking a picture shows all eight selection handles without changing its OOXML', async ({ page }, info) => {
  await open(page);
  const before = await save(page, info, 'before-selection');
  await pictures(page).first().click();
  await expect(handles(page)).toHaveCount(8);
  for (const handle of ['nw', 'ne', 'se', 'sw', 'n', 's', 'e', 'w']) {
    await expect(page.locator(`.canvas-image-selection-overlay [data-handle="${handle}"]`)).toBeVisible();
  }
  await expectPainted(page);
  const selected = await save(page, info, 'selected');
  expect(selected.state).toEqual(before.state);
  expect(await selected.zip.file('word/document.xml')!.async('string')).toBe(await before.zip.file('word/document.xml')!.async('string'));
  await pictures(page).first().click();
  await page.keyboard.press('Escape');
  await expect(handles(page)).toHaveCount(0);
  await pictures(page).first().click();
  await expect(handles(page)).toHaveCount(8);
  const destination = (await mirror(page).locator('.layout-paragraph').filter({ hasText: /^Destination$/ }).locator('.layout-run-text').first().boundingBox())!;
  await page.mouse.click(destination.x + destination.width / 2, destination.y + destination.height / 2);
  await expect(handles(page)).toHaveCount(0);
  await reopen(page, info, before);
});

test('dragging a corner resizes a picture with its aspect ratio and one-step undo', async ({ page }, info) => {
  await open(page);
  const before = await save(page, info, 'before-resize');
  const original = await imageBounds(page);
  await pictures(page).first().click();
  const corner = page.locator('.canvas-image-selection-overlay [data-handle="se"]');
  await expect(corner).toBeVisible();
  const box = (await corner.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 60, box.y + box.height / 2 + 10, { steps: 5 });
  await expect(page.locator('.paged-editor__picture-resize-preview')).toBeVisible();
  await page.mouse.up();
  await expect(page.locator('.paged-editor__picture-resize-preview')).toHaveCount(0);
  await expect.poll(async () => (await imageBounds(page)).width).toBeGreaterThan(original.width + 20);
  const resizedBox = await imageBounds(page);
  expect(resizedBox.width / resizedBox.height).toBeCloseTo(2, 1);
  const resized = await save(page, info, 'resized');
  expect(resized.state.drawings[0]!.width).toBeGreaterThan(before.state.drawings[0]!.width);
  expect(resized.state.drawings[0]!.width / resized.state.drawings[0]!.height).toBeCloseTo(2, 2);
  const redone = await history(page, info, before, resized,
    async () => { await expect.poll(async () => (await imageBounds(page)).width).toBeCloseTo(original.width, 0); },
    async () => { await expect.poll(async () => (await imageBounds(page)).width).toBeCloseTo(resizedBox.width, 0); });
  await reopen(page, info, redone);
  await expectPainted(page);
});

const dragPreview = (page: Page) => page.locator('.paged-editor__picture-drag-preview');

test('an east handle changes only width and persists after save and reopen', async ({ page }, info) => {
  await open(page);
  const before = await save(page, info, 'before-side-resize');
  const original = await imageBounds(page);
  await pictures(page).first().click();
  const side = page.locator('.canvas-image-selection-overlay [data-handle="e"]');
  const box = (await side.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 60, box.y + box.height / 2 + 25, { steps: 5 });
  await page.mouse.up();
  await expect.poll(async () => (await imageBounds(page)).width).toBeGreaterThan(original.width + 20);
  expect((await imageBounds(page)).height).toBeCloseTo(original.height, 1);
  const resized = await save(page, info, 'side-resized');
  expect(resized.state.drawings[0]!.width).toBeGreaterThan(before.state.drawings[0]!.width);
  expect(resized.state.drawings[0]!.height).toBe(before.state.drawings[0]!.height);
  const reopened = await reopen(page, info, resized);
  expect(reopened.state).toEqual(resized.state);
  await expect.poll(async () => (await imageBounds(page)).width).toBeGreaterThan(original.width + 20);
  expect((await imageBounds(page)).height).toBeCloseTo(original.height, 1);
});

test('Escape cancels a corner resize without changing the saved document', async ({ page }, info) => {
  await open(page);
  const before = await save(page, info, 'before-corner-cancel');
  const original = await imageBounds(page);
  await pictures(page).first().click();
  const corner = page.locator('.canvas-image-selection-overlay [data-handle="nw"]');
  const box = (await corner.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x - 60, box.y - 30, { steps: 5 });
  await expect(page.locator('.paged-editor__picture-resize-preview')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('.paged-editor__picture-resize-preview')).toHaveCount(0);
  await page.mouse.move(box.x - 90, box.y - 45, { steps: 3 });
  await page.mouse.up();
  expect(await imageBounds(page)).toEqual(original);
  expect((await save(page, info, 'corner-cancelled')).state).toEqual(before.state);
});

for (const kind of ['inline', 'floating', 'behind'] as const) {
  // Moving an inline picture to the drop text position needs a worker op that relocates the embed.
  const define = kind === 'inline' ? test.fixme : test;
  define(`dragging ${kind === 'inline' ? 'an inline' : kind === 'floating' ? 'a floating' : 'a behind-text'} picture moves it and supports undo, redo, and reopen`, async ({ page }, info) => {
    await open(page, undefined, kind);
    const before = await save(page, info, 'before-drag');
    const original = await imageBounds(page);
    await pictures(page).first().click();
    if (kind === 'floating') await expect(handles(page)).toHaveCount(8);
    const destination = mirror(page).locator('.layout-paragraph').filter({ hasText: /^Destination$/ }).locator('.layout-run-text').last();
    await expect(destination).toBeAttached();
    const target = (await destination.boundingBox())!;
    const end = kind === 'inline'
      ? { x: target.x + target.width + 2, y: target.y + target.height / 2 }
      : { x: original.x + original.width / 2 + 90, y: original.y + original.height / 2 + 50 };
    await page.mouse.move(original.x + original.width / 2, original.y + original.height / 2);
    await page.mouse.down();
    await page.mouse.move(end.x, end.y, { steps: 6 });
    if (kind !== 'inline') await expect(dragPreview(page)).toBeVisible();
    await page.mouse.up();
    await expect(dragPreview(page)).toHaveCount(0);
    await expect.poll(async () => Math.abs((await imageBounds(page)).y - original.y)).toBeGreaterThan(10);
    await expectPainted(page);
    const moved = await save(page, info, 'moved');
    expect(moved.state.drawings).toHaveLength(1);
    if (kind === 'inline') {
      expect(moved.state.drawings[0]!.kind).toBe('inline');
      expect(moved.state.paragraphs).toEqual(['Tail', 'Destination\uFFFC']);
    } else {
      expect(moved.state.drawings[0]!.kind).toBe('anchor');
      expect(moved.state.drawings[0]!.offsets[0]).toBeGreaterThan(before.state.drawings[0]!.offsets[0]!);
      expect(moved.state.drawings[0]!.offsets[1]).toBeGreaterThan(before.state.drawings[0]!.offsets[1]!);
      expect(moved.state.drawings[0]).toMatchObject({ wrap: before.state.drawings[0]!.wrap, behind: before.state.drawings[0]!.behind });
    }
    const movedBox = await imageBounds(page);
    const redone = await history(page, info, before, moved,
      async () => { await expect.poll(async () => (await imageBounds(page)).y).toBeCloseTo(original.y, 0); },
      async () => { await expect.poll(async () => (await imageBounds(page)).y).toBeCloseTo(movedBox.y, 0); });
    await reopen(page, info, redone);
    await expectPainted(page);
  });
}

test('Escape cancels a floating picture drag without changing the document', async ({ page }, info) => {
  await open(page, undefined, 'floating');
  const before = await save(page, info, 'before-drag');
  const original = await imageBounds(page);
  await pictures(page).first().click();
  await page.mouse.move(original.x + original.width / 2, original.y + original.height / 2);
  await page.mouse.down();
  await page.mouse.move(original.x + original.width / 2 + 90, original.y + original.height / 2 + 50, { steps: 6 });
  await expect(dragPreview(page)).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(dragPreview(page)).toHaveCount(0);
  await page.mouse.move(original.x + original.width / 2 + 120, original.y + original.height / 2 + 80, { steps: 3 });
  await page.mouse.up();
  await expect(dragPreview(page)).toHaveCount(0);
  expect(await imageBounds(page)).toEqual(original);
  expect((await save(page, info, 'cancelled')).state).toEqual(before.state);
});

test('dragging across text still selects it when the press is not on a picture', async ({ page }, info) => {
  await open(page, undefined, 'floating');
  const before = await save(page, info, 'before-select');
  const runs = mirror(page).locator('.layout-paragraph').filter({ hasText: /^Destination$/ }).locator('.layout-run-text');
  await expect(runs.first()).toBeAttached();
  const first = (await runs.first().boundingBox())!;
  const last = (await runs.last().boundingBox())!;
  await page.mouse.move(first.x + 1, first.y + first.height / 2);
  await page.mouse.down();
  await page.mouse.move(last.x + last.width - 1, last.y + last.height / 2, { steps: 6 });
  await expect(dragPreview(page)).toHaveCount(0);
  await page.mouse.up();
  await input(page).press('Backspace');
  const edited = await save(page, info, 'selection-deleted');
  expect(edited.state.paragraphs[0]).toBe('\uFFFCTail');
  expect(edited.state.paragraphs[1]!.length).toBeLessThan(4);
  expect(edited.state.drawings).toEqual(before.state.drawings);
});

for (const key of ['Delete', 'Backspace']) {
  test(`${key} removes only the selected picture and undo restores its pixels and media`, async ({ page }, info) => {
    await open(page);
    const source = (await raster(page)).buffer;
    const before = await save(page, info, 'before-delete');
    await selectByKeyboard(page);
    await input(page).press(key);
    await expect(pictures(page)).toHaveCount(0);
    await expect(mirror(page)).toContainText('Tail');
    const deleted = await save(page, info, 'deleted');
    expect(deleted.state.drawings).toEqual([]);
    expect(deleted.state.paragraphs).toEqual(['Tail', 'Destination']);
    await input(page).focus();
    await input(page).press('ControlOrMeta+z');
    await expectPainted(page);
    const restored = await save(page, info, 'restored');
    expect(restored.state).toEqual(before.state);
    await expectMedia(page, restored, source);
    await input(page).focus();
    await input(page).press('ControlOrMeta+Shift+z');
    await expect(pictures(page)).toHaveCount(0);
    const redone = await save(page, info, 'deleted-again');
    expect(redone.state).toEqual(deleted.state);
    await reopen(page, info, redone);
    await expect(mirror(page)).toContainText('Tail');
  });
}

// The worker clipboard slice carries only text and tables, so a picture-only selection copies nothing yet.
test.fixme('copy and paste duplicates a picture with an independent drawing and valid media', async ({ page, context }, info) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await open(page);
  const source = (await raster(page)).buffer;
  const before = await save(page, info, 'before-copy');
  await page.evaluate(() => navigator.clipboard.writeText('clipboard sentinel'));
  await selectByKeyboard(page);
  await input(page).press('ControlOrMeta+c');
  await expect.poll(() => page.evaluate(async () => {
    const items = await navigator.clipboard.read();
    if (items.some(item => item.types.some(type => type.startsWith('image/')))) return true;
    const html = items.find(item => item.types.includes('text/html'));
    if (!html) return false;
    const content = await (await html.getType('text/html')).text();
    return content.includes('<img') || content.includes('green-picture');
  })).toBe(true);
  await input(page).press('ControlOrMeta+End');
  await input(page).press('ControlOrMeta+v');
  await expect(pictures(page)).toHaveCount(2);
  await expectPainted(page, pictures(page).nth(1));
  const pasted = await save(page, info, 'pasted');
  expect(pasted.state.drawings).toHaveLength(2);
  expect(pasted.state.paragraphs).toEqual(['\uFFFCTail', 'Destination\uFFFC']);
  expect(pasted.state.drawings[1]).toMatchObject({ kind: 'inline', alt: imageName, width: 120 * emu, height: 60 * emu });
  await expectMedia(page, pasted, source);
  const ids = await page.evaluate(async xml => Array.from(new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 'docPr')).map(node => node.getAttribute('id')), await pasted.zip.file('word/document.xml')!.async('string'));
  expect(new Set(ids).size).toBe(2);
  const redone = await history(page, info, before, pasted,
    () => expect(pictures(page)).toHaveCount(1), () => expect(pictures(page)).toHaveCount(2));
  const reopened = await reopen(page, info, redone);
  await expectMedia(page, reopened, source);
  await expectPainted(page, pictures(page).nth(1));
});

test('image alt text can be cancelled, applied, cleared, and restored with history', async ({ page }, info) => {
  await open(page);
  const before = await save(page, info, 'before-alt');
  await selectByKeyboard(page);
  let dialog = await properties(page);
  await expect(dialog.locator('textarea')).toHaveValue(imageName);
  await dialog.locator('textarea').fill('Discarded description');
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  expect((await save(page, info, 'cancelled')).state).toEqual(before.state);
  await selectByKeyboard(page);
  dialog = await properties(page);
  await expect(dialog.locator('textarea')).toHaveValue(imageName);
  await dialog.locator('textarea').fill('Green landscape');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(pictures(page).first()).toHaveAttribute('aria-label', 'Green landscape');
  const described = await save(page, info, 'described');
  expect(described.state.drawings[0]!.alt).toBe('Green landscape');
  await history(page, info, before, described,
    () => expect(pictures(page).first()).toHaveAttribute('aria-label', imageName),
    () => expect(pictures(page).first()).toHaveAttribute('aria-label', 'Green landscape'));
  await selectByKeyboard(page);
  dialog = await properties(page);
  await dialog.locator('textarea').fill('');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(pictures(page).first()).not.toHaveAttribute('aria-label', 'Green landscape');
  const cleared = await save(page, info, 'cleared');
  expect(cleared.state.drawings[0]!.alt).toBe('');
  const redone = await history(page, info, described, cleared,
    () => expect(pictures(page).first()).toHaveAttribute('aria-label', 'Green landscape'),
    () => expect(pictures(page).first()).not.toHaveAttribute('aria-label', 'Green landscape'));
  await reopen(page, info, redone);
  await expectPainted(page);
});

test('image properties lock aspect ratio and allow independent dimensions when unlocked', async ({ page }, info) => {
  await open(page);
  const before = await save(page, info, 'before-dimensions');
  const original = await imageBounds(page);
  await selectByKeyboard(page);
  let dialog = await properties(page);
  await expect(dialog.getByRole('checkbox', { name: 'Lock aspect ratio', exact: true })).toBeChecked();
  await dialog.locator('input[type=number]').nth(0).fill('180');
  await expect(dialog.locator('input[type=number]').nth(1)).toHaveValue('90');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect.poll(async () => (await imageBounds(page)).width).toBeCloseTo(original.width * 1.5, 0);
  await expect.poll(async () => (await imageBounds(page)).height).toBeCloseTo(original.height * 1.5, 0);
  const locked = await save(page, info, 'locked-dimensions');
  expect(locked.state.drawings[0]).toMatchObject({ width: 180 * emu, height: 90 * emu });
  await history(page, info, before, locked,
    async () => { await expect.poll(async () => (await imageBounds(page)).width).toBeCloseTo(original.width, 0); },
    async () => { await expect.poll(async () => (await imageBounds(page)).width).toBeCloseTo(original.width * 1.5, 0); });
  await selectByKeyboard(page);
  dialog = await properties(page);
  await expect(dialog.locator('input[type=number]').nth(0)).toHaveValue('180');
  await dialog.getByRole('checkbox', { name: 'Lock aspect ratio', exact: true }).uncheck();
  await dialog.locator('input[type=number]').nth(1).fill('45');
  await expect(dialog.locator('input[type=number]').nth(0)).toHaveValue('180');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect.poll(async () => (await imageBounds(page)).height).toBeCloseTo(original.height * 0.75, 0);
  const unlocked = await save(page, info, 'unlocked-dimensions');
  expect(unlocked.state.drawings[0]).toMatchObject({ width: 180 * emu, height: 45 * emu });
  const redone = await history(page, info, locked, unlocked,
    async () => { await expect.poll(async () => (await imageBounds(page)).height).toBeCloseTo(original.height * 1.5, 0); },
    async () => { await expect.poll(async () => (await imageBounds(page)).height).toBeCloseTo(original.height * 0.75, 0); });
  await reopen(page, info, redone);
  await expectPainted(page);
});

test('context wrapping changes an inline picture to square right and back with history', async ({ page }, info) => {
  await open(page);
  const before = await save(page, info, 'before-wrap');
  const original = await imageBounds(page);
  await wrap(page, 'Square Right');
  await expect.poll(async () => (await imageBounds(page)).x).toBeGreaterThan(original.x + 100);
  await expectPainted(page);
  const wrapped = await save(page, info, 'square-right');
  expect(wrapped.state.drawings[0]).toMatchObject({ kind: 'anchor', wrap: 'wrapSquare', side: 'left', behind: '0' });
  await history(page, info, before, wrapped,
    async () => { await expect.poll(async () => (await imageBounds(page)).x).toBeCloseTo(original.x, 0); },
    async () => { await expect.poll(async () => (await imageBounds(page)).x).toBeGreaterThan(original.x + 100); });
  await wrap(page, 'In Line with Text');
  await expect.poll(async () => (await imageBounds(page)).x).toBeCloseTo(original.x, 0);
  const inline = await save(page, info, 'inline-again');
  expect(inline.state.drawings[0]).toMatchObject({ kind: 'inline', wrap: '' });
  expect(inline.state.paragraphs).toEqual(before.state.paragraphs);
  const redone = await history(page, info, wrapped, inline,
    async () => { await expect.poll(async () => (await imageBounds(page)).x).toBeGreaterThan(original.x + 100); },
    async () => { await expect.poll(async () => (await imageBounds(page)).x).toBeCloseTo(original.x, 0); });
  await reopen(page, info, redone);
  await expectPainted(page);
});

for (const kind of ['textbox', 'shape'] as const) {
  // A click in text box text still places the worker caret in the body story, so typing lands in the body.
  test.fixme(`typing inside a ${kind === 'textbox' ? 'text box' : 'shape'} edits txbxContent without changing body text`, async ({ page }, info) => {
    await open(page, undefined, kind);
    const before = await save(page, info, 'before-box-edit');
    const text = mirror(page).locator('.layout-paragraph').filter({ hasText: /^Box seed$/ }).last().locator('.layout-run-text').first();
    await expect(text).toBeAttached();
    const bounds = (await text.boundingBox())!;
    await page.mouse.dblclick(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
    await input(page).press('End');
    await page.keyboard.insertText(' edited');
    await expect(mirror(page)).toContainText('Box seed edited');
    const edited = await save(page, info, 'box-edited');
    expect(edited.state.boxes).toEqual(['Box seed edited']);
    expect(edited.state.paragraphs).toEqual(before.state.paragraphs);
    expect(edited.state.drawings).toEqual(before.state.drawings);
    const redone = await history(page, info, before, edited,
      async () => {
        await expect(mirror(page)).toContainText('Box seed');
        await expect(mirror(page)).not.toContainText(' edited');
      }, () => expect(mirror(page)).toContainText('Box seed edited'));
    await reopen(page, info, redone);
    await expect(mirror(page)).toContainText('Box seed edited');
  });
}

test('northwest floating resize keeps the opposite corner through save, undo and reopen', async ({ page }, info) => {
  await open(page, undefined, 'floating');
  const before = await save(page, info, 'before-northwest');
  const original = await imageBounds(page);
  await pictures(page).first().click();
  const handle = page.locator('.canvas-image-selection-overlay [data-handle="nw"]');
  await expect(handle).toBeVisible();
  const box = (await handle.boundingBox())!;
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + 30, y, { steps: 5 });
  const preview = page.locator('.paged-editor__picture-resize-preview');
  await expect(preview).toBeVisible();
  const expected = (await preview.boundingBox())!;
  expect(expected.width).toBeLessThan(original.width - 10);
  await page.mouse.up();
  await expect.poll(async () => (await imageBounds(page)).x).toBeCloseTo(expected.x, 0);
  const resizedBox = await imageBounds(page);
  expect(resizedBox.y).toBeCloseTo(expected.y, 0);
  expect(resizedBox.width).toBeCloseTo(expected.width, 0);
  expect(resizedBox.height).toBeCloseTo(expected.height, 0);
  expect(resizedBox.x + resizedBox.width).toBeCloseTo(original.x + original.width, 0);
  expect(resizedBox.y + resizedBox.height).toBeCloseTo(original.y + original.height, 0);
  const resized = await save(page, info, 'northwest-resized');
  const redone = await history(page, info, before, resized,
    async () => { await expect.poll(async () => (await imageBounds(page)).x).toBeCloseTo(original.x, 0); },
    async () => { await expect.poll(async () => (await imageBounds(page)).x).toBeCloseTo(resizedBox.x, 0); });
  await reopen(page, info, redone);
  await expect.poll(async () => (await imageBounds(page)).width).toBeCloseTo(resizedBox.width, 0);
});

test('rotated pictures keep selection but do not offer resize handles', async ({ page }, info) => {
  const file = await fixture(page);
  const zip = await JSZip.loadAsync(file.buffer);
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<a:xfrm>', '<a:xfrm rot="5400000">'));
  file.buffer = await zip.generateAsync({ type: 'nodebuffer' });
  await open(page, file);
  await selectByKeyboard(page);
  await expect(page.locator('.canvas-image-selection-overlay')).toBeVisible();
  await expect(handles(page)).toHaveCount(0);
  const saved = await save(page, info, 'rotated-selected');
  expect(await saved.zip.file('word/document.xml')!.async('string')).toBe(await zip.file('word/document.xml')!.async('string'));
});
