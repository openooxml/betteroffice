import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const drawing = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const presentation = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const namespaces = `xmlns:p="${presentation}" xmlns:a="${drawing}" xmlns:r="${relationship}"`;
const emu = (value: number) => Math.round(value * 9525);
const canvas = (page: Page) => page.getByTestId('pptx-slide-canvas');
const input = (page: Page) => page.getByRole('textbox', { name: 'Slide text', exact: true });

function cell(text: string) {
  const run = text ? `<a:r><a:rPr sz="2000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Liberation Sans"/></a:rPr><a:t>${text}</a:t></a:r>` : '';
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${run}<a:endParaRPr sz="2000"/></a:p></a:txBody><a:tcPr marL="95250" marR="95250" marT="95250" marB="95250"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr></a:tc>`;
}

async function fixture() {
  const rows = [['A1', 'B1'], ['A2', 'B2']].map((row) => `<a:tr h="${emu(90)}">${row.map(cell).join('')}</a:tr>`).join('');
  const table = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(100)}" y="${emu(100)}"/><a:ext cx="${emu(600)}" cy="${emu(180)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="${emu(300)}"/><a:gridCol w="${emu(300)}"/></a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
  const zip = new JSZip();
  const relationships = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', relationships(`<Relationship Id="rId1" Type="${relationship}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation ${namespaces}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="${emu(960)}" cy="${emu(540)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', relationships(`<Relationship Id="rId1" Type="${relationship}/slide" Target="slides/slide1.xml"/>`));
  const tree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
  zip.file('ppt/slides/slide1.xml', `<p:sld ${namespaces}><p:cSld><p:spTree>${tree}${table}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page) {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'table.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: await fixture(),
  });
  await expect(page.getByTestId('pptx-worker-editor')).toBeVisible();
  await expect(canvas(page)).toBeVisible();
  await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
  await expect(page.getByTestId('pptx-error')).toHaveCount(0);
}

async function editCell(page: Page, x: number, y: number) {
  const bounds = await canvas(page).boundingBox();
  if (!bounds) throw new Error('Slide canvas is missing');
  await page.mouse.click(bounds.x + x * bounds.width / 960, bounds.y + y * bounds.height / 540);
  await expect(input(page)).toBeFocused();
  await expect(page.getByTestId('pptx-text-edit-outline')).toBeVisible();
}

async function cells(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${label}.pptx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
  return page.evaluate((source) => {
    const document = new DOMParser().parseFromString(source, 'application/xml');
    return Array.from(document.getElementsByTagNameNS('*', 'tc')).map((tc) =>
      Array.from(tc.getElementsByTagNameNS('*', 't')).map((text) => text.textContent ?? '').join(''));
  }, xml);
}

test('pptx table: Tab selects the next cell, wrapping to the next row', async ({ page }, info) => {
  await open(page);
  await editCell(page, 135, 125);
  await input(page).press('Tab');
  await expect(input(page)).toBeFocused();
  await page.keyboard.insertText('x');
  await input(page).press('Tab');
  await page.keyboard.insertText('y');
  await expect.poll(() => cells(page, info, 'tab')).toEqual(['A1', 'x', 'y', 'B2']);
});

test('pptx table: Shift+Tab selects the previous cell, wrapping to the previous row', async ({ page }, info) => {
  await open(page);
  await editCell(page, 435, 215);
  await input(page).press('Shift+Tab');
  await expect(input(page)).toBeFocused();
  await page.keyboard.insertText('x');
  await input(page).press('Shift+Tab');
  await page.keyboard.insertText('y');
  await expect.poll(() => cells(page, info, 'shift-tab')).toEqual(['A1', 'y', 'x', 'B2']);
});

test('pptx table: Tab in the last cell keeps the caret there', async ({ page }, info) => {
  await open(page);
  await editCell(page, 435, 215);
  await input(page).press('End');
  await input(page).press('Tab');
  await expect(input(page)).toBeFocused();
  await page.keyboard.insertText('!');
  await expect.poll(() => cells(page, info, 'last')).toEqual(['A1', 'B1', 'A2', 'B2!']);
});
