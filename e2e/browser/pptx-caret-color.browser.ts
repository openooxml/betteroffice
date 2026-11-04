import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { inputReady, slideCanvas } from './pptx-async-helpers';

const DARK_HOST = 'rgb(237, 237, 237)';

async function deck(background: string | null, color: string) {
  const zip = new JSZip();
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(`<Relationship Id="rId1" Type="${r}/slide" Target="slides/slide1.xml"/>`));
  const bg = background ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${background}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : '';
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text 2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Caret text</a:t></a:r><a:endParaRPr sz="2400"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return { name: 'caret.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function edit(page: Page, background: string | null, color: string, host: boolean) {
  await page.goto('/?format=pptx');
  if (host) await page.addStyleTag({ content: `body { color: ${DARK_HOST}; }` });
  await page.locator('input[type=file]').last().setInputFiles(await deck(background, color));
  await inputReady(page);
  const canvas = (await slideCanvas(page).boundingBox())!;
  await page.mouse.click(canvas.x + canvas.width * 0.15, canvas.y + canvas.height * 0.22);
  const input = page.getByRole('textbox', { name: 'Slide text', exact: true });
  await expect(input).toBeFocused();
  await input.press('End');
  await expect(page.getByTestId('pptx-caret')).toBeAttached();
}

const caretColor = (page: Page) => page.getByTestId('pptx-caret').evaluate((element) => getComputedStyle(element).backgroundColor);

function luminance(color: string) {
  const [r, g, b] = color.match(/\d+(\.\d+)?/g)!.slice(0, 3).map((value) => {
    const c = Number(value) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r! + 0.7152 * g! + 0.0722 * b!;
}

function ratio(a: string, b: string) {
  const [light, dark] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (light! + 0.05) / (dark! + 0.05);
}

for (const host of [false, true]) {
  const suffix = host ? ' inside a dark-themed host' : '';
  test(`pptx worker caret takes the text colour on a white slide${suffix}`, async ({ page }) => {
    await edit(page, null, 'C00000', host);
    await expect.poll(() => caretColor(page)).toBe('rgb(192, 0, 0)');
  });

  test(`pptx worker caret stays visible with white text on a dark slide${suffix}`, async ({ page }) => {
    await edit(page, '0B1F3A', 'FFFFFF', host);
    await expect.poll(() => caretColor(page)).toBe('rgb(255, 255, 255)');
  });
}

test('pptx worker composition preedit reads against the slide inside a dark-themed host', async ({ page }) => {
  await edit(page, null, '172554', true);
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: '日本', selectionStart: 2, selectionEnd: 2 });
  const preedit = page.getByTestId('pptx-preedit');
  await expect(preedit).toBeVisible();
  const colors = await preedit.evaluate((element) => {
    const style = getComputedStyle(element);
    return { text: style.color, background: style.backgroundColor };
  });
  expect(colors.text).toBe('rgb(23, 37, 84)');
  expect(ratio(colors.text, colors.background)).toBeGreaterThanOrEqual(3);
});
