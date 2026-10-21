import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import JSZip from 'jszip';

type Format = 'docx' | 'xlsx' | 'pptx';
type InputFile = { name: string; mimeType: string; buffer: Buffer };
type BrokenFile = { title: string; file: InputFile; message: RegExp };

const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageRels = 'http://schemas.openxmlformats.org/package/2006/relationships';
const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const namespaces = {
  docx: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  xlsx: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  pptx: 'http://schemas.openxmlformats.org/presentationml/2006/main',
};
const mainParts = {
  docx: 'word/document.xml',
  xlsx: 'xl/workbook.xml',
  pptx: 'ppt/presentation.xml',
};
const contentParts = {
  docx: 'word/document.xml',
  xlsx: 'xl/worksheets/sheet1.xml',
  pptx: 'ppt/slides/slide1.xml',
};
const mimeTypes = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};
const before = 'Before object';
const after = 'After object';
const pageErrors = new WeakMap<Page, string[]>();

test.beforeEach(async ({ page }) => {
  const errors: string[] = [];
  pageErrors.set(page, errors);
  page.on('pageerror', (error) => errors.push(error.message));
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.route('**/*', async (route) => {
    const url = new URL(route.request().url());
    if (['http:', 'https:'].includes(url.protocol) && url.hostname !== '127.0.0.1') {
      errors.push('Unexpected external request: ' + url.href);
      await route.abort();
    } else await route.continue();
  });
});

test.afterEach(async ({ page }) => {
  expect.soft(pageErrors.get(page), 'Unhandled page errors and external requests').toEqual([]);
});

function relationships(entries: string) {
  return `<Relationships xmlns="${packageRels}">${entries}</Relationships>`;
}

function relationship(id: string, type: string, target: string) {
  return `<Relationship Id="${id}" Type="${office}/${type}" Target="${target}"/>`;
}

function slideText(id: number, text: string, y: number) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="914400" y="${y}"/><a:ext cx="6400800" cy="731520"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0" anchor="t"/><a:lstStyle/><a:p><a:r><a:rPr sz="2000"><a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function fixture(format: Format, first = before) {
  const zip = new JSZip();
  const overrides: [string, string][] = [[mainParts[format], mimeTypes[format] + '.main+xml']];
  zip.file('_rels/.rels', relationships(relationship('rIdOffice', 'officeDocument', mainParts[format])));
  if (format === 'docx') {
    zip.file(mainParts.docx, `<w:document xmlns:w="${namespaces.docx}" xmlns:r="${office}" xmlns:a="${drawing}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>` +
      `<w:p w14:paraId="00000001"><w:r><w:t>${first}</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>${after}</w:t></w:r></w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  } else if (format === 'xlsx') {
    overrides.push([contentParts.xlsx, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']);
    zip.file(mainParts.xlsx, `<workbook xmlns="${namespaces.xlsx}" xmlns:r="${office}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', relationships(relationship('rIdSheet', 'worksheet', 'worksheets/sheet1.xml')));
    zip.file(contentParts.xlsx, `<worksheet xmlns="${namespaces.xlsx}" xmlns:r="${office}"><dimension ref="A1:A3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="32" customWidth="1"/></cols>` +
      `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${first}</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>${after}</t></is></c></row></sheetData></worksheet>`);
  } else {
    overrides.push([contentParts.pptx, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']);
    const ns = `xmlns:p="${namespaces.pptx}" xmlns:a="${drawing}" xmlns:r="${office}"`;
    zip.file(mainParts.pptx, `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
    zip.file('ppt/_rels/presentation.xml.rels', relationships(relationship('rIdSlide', 'slide', 'slides/slide1.xml')));
    zip.file(contentParts.pptx, `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      slideText(2, first, 914400) + slideText(3, after, 2743200) + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>');
  }
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    overrides.map(([path, type]) => `<Override PartName="/${path}" ContentType="${type}"/>`).join('') + '</Types>');
  return zip;
}

async function inputFile(zip: JSZip, format: Format, name: string): Promise<InputFile> {
  return { name: name + '.' + format, mimeType: mimeTypes[format],
    buffer: await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }) };
}

async function append(zip: JSZip, part: string, closing: string, xml: string) {
  zip.file(part, (await zip.file(part)!.async('string')).replace(closing, xml + closing));
}

async function addRelationship(zip: JSZip, part: string, xml: string) {
  if (zip.file(part)) await append(zip, part, '</Relationships>', xml);
  else zip.file(part, relationships(xml));
}

async function addContent(zip: JSZip, format: Format, xml: string) {
  const closing = { docx: '<w:p w14:paraId="00000002">', xlsx: '</worksheet>', pptx: '</p:spTree>' };
  await append(zip, contentParts[format], closing[format], xml);
}

function oleHeader() {
  const bytes = Buffer.alloc(512);
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(bytes);
  bytes.writeUInt16LE(0x003e, 24);
  bytes.writeUInt16LE(3, 26);
  bytes.writeUInt16LE(0xfffe, 28);
  bytes.writeUInt16LE(9, 30);
  bytes.writeUInt16LE(6, 32);
  bytes.writeUInt32LE(0xfffffffe, 48);
  bytes.writeUInt32LE(4096, 56);
  bytes.writeUInt32LE(0xfffffffe, 60);
  bytes.writeUInt32LE(0xfffffffe, 68);
  for (let offset = 76; offset < bytes.length; offset += 4) bytes.writeUInt32LE(0xffffffff, offset);
  return bytes;
}

async function invalidArchives(format: Format): Promise<BrokenFile[]> {
  const ordinary = new JSZip();
  ordinary.file('plain.txt', 'An ordinary ZIP archive without Office parts.');
  const complete = await inputFile(fixture(format), format, 'truncated');
  const file = (name: string, buffer: Buffer): InputFile => ({ name: name + '.' + format, mimeType: mimeTypes[format], buffer });
  return [
    { title: 'ZIP without OOXML', file: await inputFile(ordinary, format, 'ordinary-zip'), message: /office|document|workbook|presentation|part|content.?types|ooxml|missing|not found/i },
    { title: 'truncated ZIP', file: file('truncated', complete.buffer.subarray(0, Math.floor(complete.buffer.length / 2))), message: /zip|archive|central|directory|end|eof|truncat|invalid/i },
    { title: 'encrypted OLE header', file: file('encrypted', oleHeader()), message: /encrypt|password|ole|zip|archive|office|format|invalid/i },
    { title: 'empty file', file: file('empty', Buffer.alloc(0)), message: /empty|file|zip|archive|header|end|eof|invalid/i },
  ];
}

async function invalidPackages(format: Format): Promise<BrokenFile[]> {
  const missing = fixture(format);
  missing.remove(mainParts[format]);
  const malformed = fixture(format);
  malformed.file(contentParts[format], (await malformed.file(contentParts[format])!.async('string')).replace(before, 'Broken <unclosed>'));
  const foreign = fixture(format);
  for (const part of new Set([mainParts[format], contentParts[format]])) {
    foreign.file(part, (await foreign.file(part)!.async('string')).replaceAll(namespaces[format], 'urn:unsupported-ooxml:main'));
  }
  return [
    { title: 'missing main part', file: await inputFile(missing, format, 'missing-main'), message: /missing|part|document|workbook|presentation|not found|exist|resolve/i },
    { title: 'malformed content XML', file: await inputFile(malformed, format, 'malformed-xml'), message: /xml|parse|malformed|unclosed|invalid|syntax|eof|end/i },
    { title: 'unsupported main namespace', file: await inputFile(foreign, format, 'foreign-namespace'), message: /namespace|unsupported|invalid|format|document|workbook|presentation|root/i },
  ];
}

function oversizedTiff() {
  const pixels = deflateSync(Buffer.alloc(8192 * 8192 / 8));
  const tags = [
    [256, 4, 8192], [257, 4, 8192], [258, 3, 1], [259, 3, 8], [262, 3, 1],
    [273, 4, 0], [277, 3, 1], [278, 4, 8192], [279, 4, pixels.length], [284, 3, 1],
  ];
  const header = Buffer.alloc(8 + 2 + tags.length * 12 + 4);
  header.write('II');
  header.writeUInt16LE(42, 2);
  header.writeUInt32LE(8, 4);
  header.writeUInt16LE(tags.length, 8);
  for (const [index, [tag, type, value]] of tags.entries()) {
    const offset = 10 + index * 12;
    header.writeUInt16LE(tag, offset);
    header.writeUInt16LE(type, offset + 2);
    header.writeUInt32LE(1, offset + 4);
    header.writeUInt32LE(tag === 273 ? header.length : value, offset + 8);
  }
  return Buffer.concat([header, pixels]);
}

function pictureProperties() {
  return `<a:xfrm><a:off x="5486400" y="1828800"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`;
}

async function oversizedImage(format: Format) {
  const zip = fixture(format);
  const prefix = { docx: 'word', xlsx: 'xl', pptx: 'ppt' }[format];
  const part = prefix + '/media/large.tiff';
  const bytes = oversizedTiff();
  zip.file(part, bytes);
  await append(zip, '[Content_Types].xml', '</Types>', '<Default Extension="tiff" ContentType="image/tiff"/>');
  if (format === 'docx') {
    await addRelationship(zip, 'word/_rels/document.xml.rels', relationship('rIdImage', 'image', 'media/large.tiff'));
    await addContent(zip, format, '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="4" name="Large image"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="4" name="Large image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>' + pictureProperties() + '</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>');
  } else if (format === 'xlsx') {
    await addContent(zip, format, '<drawing r:id="rIdDrawing"/>');
    await addRelationship(zip, 'xl/worksheets/_rels/sheet1.xml.rels', relationship('rIdDrawing', 'drawing', '../drawings/drawing1.xml'));
    await append(zip, '[Content_Types].xml', '</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
    zip.file('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${drawing}" xmlns:r="${office}"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="4" name="Large image"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${pictureProperties()}</xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
    zip.file('xl/drawings/_rels/drawing1.xml.rels', relationships(relationship('rIdImage', 'image', '../media/large.tiff')));
  } else {
    await addRelationship(zip, 'ppt/slides/_rels/slide1.xml.rels', relationship('rIdImage', 'image', '../media/large.tiff'));
    await addContent(zip, format, '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Large image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>' + pictureProperties() + '</p:spPr></p:pic>');
  }
  return { file: await inputFile(zip, format, 'oversized-image'), part, bytes };
}

async function unsupportedObject(format: Format) {
  const zip = fixture(format);
  const prefix = { docx: 'word', xlsx: 'xl', pptx: 'ppt' }[format];
  const part = prefix + '/embeddings/embeddedWorkbook.xlsx';
  const bytes = (await inputFile(fixture('xlsx', 'Embedded payload'), 'xlsx', 'embedded')).buffer;
  zip.file(part, bytes);
  await append(zip, '[Content_Types].xml', '</Types>', `<Override PartName="/${part}" ContentType="${mimeTypes.xlsx}"/>`);
  const rels = { docx: 'word/_rels/document.xml.rels', xlsx: 'xl/worksheets/_rels/sheet1.xml.rels', pptx: 'ppt/slides/_rels/slide1.xml.rels' }[format];
  const target = (format === 'docx' ? '' : '../') + 'embeddings/embeddedWorkbook.xlsx';
  await addRelationship(zip, rels, relationship('rIdEmbedded', 'package', target));
  if (format === 'docx') {
    await addContent(zip, format, '<w:p><w:r><w:object><v:shape id="_x0000_i1025" style="width:72pt;height:72pt"/><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1" r:id="rIdEmbedded"/></w:object></w:r></w:p>');
  } else if (format === 'xlsx') {
    await addContent(zip, format, '<oleObjects><oleObject progId="Excel.Sheet.12" shapeId="1025" r:id="rIdEmbedded"/></oleObjects>');
  } else {
    await addContent(zip, format, '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Embedded workbook"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="5486400" y="1828800"/><a:ext cx="914400" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Embedded workbook" progId="Excel.Sheet.12" r:id="rIdEmbedded" imgW="914400" imgH="914400"><p:embed/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>');
  }
  return { file: await inputFile(zip, format, 'unsupported-object'), part, bytes, rels, target };
}

async function upload(page: Page, file: InputFile | string) {
  await page.locator('input[type=file]').last().setInputFiles(file);
}

async function assertReady(page: Page, name: string) {
  await expect(page.locator('.document-name')).toHaveText(name);
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(page.locator('[data-testid$="-error"]')).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Save', exact: true }).first()).toBeEnabled();
}

const docxInput = (page: Page) => page.locator('.paged-editor__peerless-input');
const slideInput = (page: Page) => page.getByTestId('pptx-peerless-editor')
  .getByRole('textbox', { name: 'Slide text', exact: true });

async function slideLine(page: Page, last = false) {
  const input = slideInput(page);
  await expect(input).toBeEditable();
  const canvas = page.getByTestId('pptx-slide-canvas');
  const bounds = await canvas.boundingBox();
  expect(bounds).not.toBeNull();
  // The fixture's text starts at one inch and three inches on a 10 by 5.625 inch slide.
  await canvas.click({ position: { x: bounds!.width * 0.14, y: bounds!.height * (last ? 0.56 : 0.2) } });
  await expect(page.getByTestId('pptx-text-edit-outline')).toBeVisible();
  await expect(input).toBeFocused();
  await input.press('Home');
  await input.press('Shift+End');
  await page.evaluate(() => navigator.clipboard.writeText(''));
  await input.press('ControlOrMeta+c');
  let text = '';
  await expect.poll(async () => (text = await page.evaluate(() => navigator.clipboard.readText()))).not.toBe('');
  return text;
}

async function assertText(page: Page, format: Format, first: string) {
  if (format === 'docx') {
    const mirror = page.locator('.canvas-page-mirror').first();
    await expect(mirror).toContainText(first);
    await expect(mirror).toContainText(after);
  } else if (format === 'xlsx') {
    await expect(page.getByRole('gridcell').filter({ hasText: first })).toHaveCount(1);
    await expect(page.getByRole('gridcell').filter({ hasText: after })).toHaveCount(1);
  } else {
    expect(await slideLine(page)).toBe(first);
    expect(await slideLine(page, true)).toBe(after);
  }
}

async function edit(page: Page, format: Format, marker: string) {
  if (format === 'docx') {
    const input = docxInput(page);
    await input.focus();
    await input.press('ControlOrMeta+Home');
    await page.keyboard.insertText(marker);
  } else if (format === 'xlsx') {
    const grid = page.getByTestId('xlsx-scroll');
    await grid.press('ControlOrMeta+Home');
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await grid.press('F2');
    const input = page.getByTestId('xlsx-cell-editor');
    await expect(input).toHaveValue(before);
    await input.press('Home');
    await page.keyboard.insertText(marker);
    await input.press('Enter');
  } else {
    expect(await slideLine(page)).toBe(before);
    await slideInput(page).press('Home');
    await page.keyboard.insertText(marker);
  }
  await assertText(page, format, marker + before);
}

async function history(page: Page, format: Format, action: 'undo' | 'redo') {
  if (format === 'docx') {
    await docxInput(page).press(action === 'undo' ? 'ControlOrMeta+z' : 'ControlOrMeta+Shift+z');
  } else {
    const button = page.getByTestId(format + '-' + action);
    await expect(button).toBeEnabled();
    await button.click();
  }
}

async function save(page: Page, format: Format, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(label + '.' + format);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  expect(zip.file(mainParts[format])).not.toBeNull();
  const xml = await Promise.all(Object.values(zip.files)
    .filter((part) => !part.dir && /\.(xml|rels)$/.test(part.name))
    .map((part) => part.async('string')));
  await page.evaluate((parts) => {
    for (const xml of parts) {
      if (new DOMParser().parseFromString(xml, 'application/xml').querySelector('parsererror')) {
        throw new Error('Invalid saved XML');
      }
    }
  }, xml);
  const content = await zip.file(contentParts[format])!.async('string');
  const sharedStrings = await zip.file('xl/sharedStrings.xml')?.async('string');
  const text = await page.evaluate(({ xml, sharedStrings, format }) => {
    const parser = new DOMParser();
    const document = parser.parseFromString(xml, 'application/xml');
    const strings = (node: Element | Document) => Array.from(node.getElementsByTagNameNS('*', 't'))
      .map((node) => node.textContent ?? '').join('');
    if (format === 'xlsx') {
      const shared = sharedStrings ? Array.from(parser.parseFromString(sharedStrings, 'application/xml')
        .getElementsByTagNameNS('*', 'si')).map(strings) : [];
      return Array.from(document.getElementsByTagNameNS('*', 'c')).map((cell) => {
        if (cell.getAttribute('t') === 's') {
          return shared[Number(cell.getElementsByTagNameNS('*', 'v')[0]?.textContent)];
        }
        return strings(cell);
      }).join('\n');
    }
    return Array.from(document.getElementsByTagNameNS('*', 'p')).map(strings).filter(Boolean).join('\n');
  }, { xml: content, sharedStrings, format });
  await expect(page.locator('.document-name')).not.toContainText(' · Edited');
  return { file, name: label + '.' + format, zip, text, content };
}

async function rejectAndRecover(page: Page, format: Format, info: TestInfo, broken: BrokenFile, index: number) {
  const original = await inputFile(fixture(format), format, 'original-' + index);
  await upload(page, original);
  await assertReady(page, original.name);
  await assertText(page, format, before);
  await upload(page, broken.file);
  const notice = page.locator('.notice');
  await expect(notice).toBeVisible();
  await expect(notice).toHaveAttribute('role', 'alert');
  await expect(notice).toContainText(broken.message);
  await expect(notice).not.toContainText(/undefined|null|\[object Object\]/);
  await expect(page.locator('.document-name')).toHaveText(original.name);
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await notice.getByRole('button', { name: 'Dismiss error' }).click();
  await expect(notice).toHaveCount(0);

  const marker = 'Recovered' + index;
  await edit(page, format, marker);
  await history(page, format, 'undo');
  await assertText(page, format, before);
  expect((await save(page, format, info, 'undone-' + index)).text).toBe(before + '\n' + after);
  await history(page, format, 'redo');
  await assertText(page, format, marker + before);
  const edited = await save(page, format, info, 'after-rejection-' + index);
  expect(edited.text).toBe(marker + before + '\n' + after);
  await upload(page, edited.file);
  await assertReady(page, edited.name);
  await assertText(page, format, marker + before);
  expect((await save(page, format, info, 'reopened-' + index)).text).toBe(edited.text);

  const valid = await inputFile(fixture(format, 'Later valid document'), format, 'later-valid-' + index);
  await upload(page, valid);
  await assertReady(page, valid.name);
  await assertText(page, format, 'Later valid document');
  expect((await save(page, format, info, 'later-valid-' + index)).text).toBe('Later valid document\n' + after);
}

async function preserveThroughEdits(page: Page, format: Format, info: TestInfo,
  payload: { file: InputFile; part: string; bytes: Buffer },
  verify?: (saved: Awaited<ReturnType<typeof save>>) => Promise<void>) {
  await page.goto('/?format=' + format);
  await upload(page, payload.file);
  await assertReady(page, payload.file.name);
  await assertText(page, format, before);
  const preserved = async (label: string, first: string) => {
    const saved = await save(page, format, info, label);
    expect(saved.text).toBe(first + '\n' + after);
    expect(saved.zip.file(payload.part)).not.toBeNull();
    expect(await saved.zip.file(payload.part)!.async('nodebuffer')).toEqual(payload.bytes);
    await verify?.(saved);
    return saved;
  };
  await edit(page, format, 'Kept');
  await preserved('edited', 'Kept' + before);
  await history(page, format, 'undo');
  await assertText(page, format, before);
  await preserved('undone', before);
  await history(page, format, 'redo');
  await assertText(page, format, 'Kept' + before);
  const redone = await preserved('redone', 'Kept' + before);
  await upload(page, redone.file);
  await assertReady(page, redone.name);
  await assertText(page, format, 'Kept' + before);
  await preserved('reopened', 'Kept' + before);
  const valid = await inputFile(fixture(format, 'Later valid document'), format, 'later-valid');
  await upload(page, valid);
  await assertReady(page, valid.name);
  await assertText(page, format, 'Later valid document');
  expect((await save(page, format, info, 'later-valid')).text).toBe('Later valid document\n' + after);
}

for (const format of ['docx', 'xlsx', 'pptx'] as const) {
  for (const group of [
    { create: invalidArchives, titles: ['ZIP without OOXML', 'truncated ZIP', 'encrypted OLE header', 'empty file'] },
    { create: invalidPackages, titles: ['missing main part', 'malformed content XML', 'unsupported main namespace'] },
  ]) {
    for (const [index, title] of group.titles.entries()) {
      test(`${format}: rejects ${title} without losing the document`, async ({ page }, info) => {
        await page.goto('/?format=' + format);
        const files = await group.create(format);
        await rejectAndRecover(page, format, info, files[index]!, index);
      });
    }
  }

  test(`${format}: rejects another format renamed to the current extension and recovers`, async ({ page }, info) => {
    await page.goto('/?format=' + format);
    const other = format === 'docx' ? 'xlsx' : format === 'xlsx' ? 'pptx' : 'docx';
    const wrong = await inputFile(fixture(other, 'Wrong format content'), other, 'wrong-content');
    await rejectAndRecover(page, format, info, {
      title: 'mismatched extension', file: { ...wrong, name: 'wrong-content.' + format, mimeType: mimeTypes[format] },
      message: /format|office|document|workbook|presentation|content.?type|invalid|unsupported|part/i,
    }, 0);
  });

  test(`${format}: preserves an over-budget embedded image through edit, undo, redo, and reopen`, async ({ page }, info) => {
    await preserveThroughEdits(page, format, info, await oversizedImage(format));
  });

  test(`${format}: preserves an unsupported embedded workbook through edit, undo, redo, and reopen`, async ({ page }, info) => {
    const object = await unsupportedObject(format);
    await preserveThroughEdits(page, format, info, object, async (saved) => {
      expect(saved.content).toContain('rIdEmbedded');
      const rels = await saved.zip.file(object.rels)!.async('string');
      const link = await page.evaluate(({ xml, namespace }) => {
        const node = Array.from(new DOMParser().parseFromString(xml, 'application/xml')
          .getElementsByTagNameNS(namespace, 'Relationship')).find((node) => node.getAttribute('Id') === 'rIdEmbedded');
        return node ? { target: node.getAttribute('Target'), type: node.getAttribute('Type') } : null;
      }, { xml: rels, namespace: packageRels });
      expect(link).toEqual({ target: object.target, type: office + '/package' });
    });
  });
}
