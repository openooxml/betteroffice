import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { crc32, deflateSync } from 'node:zlib';
import { inputReady, slideCanvas } from './pptx-async-helpers';

const DARK_HOST = 'rgb(237, 237, 237)';
const BLACK = 'rgb(0, 0, 0)';
const WHITE = 'rgb(255, 255, 255)';
const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const emu = (px: number) => Math.round(px * 9525);

interface Box { x: number; y: number; w: number; h: number }
const BOX: Box = { x: 96, y: 96, w: 384, h: 96 };

const solid = (color: string) => `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
const xfrm = (box: Box, rot = 0) => `<a:xfrm${rot ? ` rot="${rot * 60000}"` : ''}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emu(box.w)}" cy="${emu(box.h)}"/></a:xfrm>`;
const run = (color: string, text: string, size = 2400) =>
  `<a:r><a:rPr sz="${size}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r><a:endParaRPr sz="${size}"/>`;

function shape(color: string, options: { box?: Box; fill?: string; prst?: string; rot?: number; line?: string; algn?: string } = {}) {
  const line = options.line ? `<a:ln w="12700">${solid(options.line)}</a:ln>` : '';
  return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Text 2"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(options.box ?? BOX, options.rot)}<a:prstGeom prst="${options.prst ?? 'rect'}"><a:avLst/></a:prstGeom>${options.fill ?? '<a:noFill/>'}${line}</p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p>${options.algn ? `<a:pPr algn="${options.algn}"/>` : ''}${run(color, 'Caret text')}</a:p></p:txBody></p:sp>`;
}

function png(width: number, pixel: (x: number) => number) {
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
    const length = Buffer.alloc(4); length.writeUInt32BE(data.length);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(body));
    return Buffer.concat([length, body, crc]);
  };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0); header.writeUInt32BE(1, 4); header[8] = 8; header[9] = 2;
  const row = Buffer.alloc(1 + width * 3);
  for (let x = 0; x < width; x += 1) row.fill(pixel(x), 1 + x * 3, 4 + x * 3);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header),
    chunk('IDAT', deflateSync(row)), chunk('IEND', Buffer.alloc(0))]);
}

const photo = `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm({ x: 0, y: 0, w: 960, h: 540 })}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

function table(cells: readonly (readonly [fill: string, color: string])[]) {
  const cell = ([fill, color]: readonly [string, string]) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${run(color, 'Type', 2800)}</a:p></a:txBody><a:tcPr marL="95250" marR="95250" marT="95250" marB="95250">${solid(fill)}</a:tcPr></a:tc>`;
  const rows = [cells.slice(0, 2), cells.slice(2, 4)].map((row) => `<a:tr h="${emu(90)}">${row.map(cell).join('')}</a:tr>`).join('');
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(96)}" y="${emu(96)}"/><a:ext cx="${emu(600)}" cy="${emu(180)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="${emu(300)}"/><a:gridCol w="${emu(300)}"/></a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

async function deck(background: string | null, body: string, image?: Buffer) {
  const zip = new JSZip();
  const rels = (inner: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(`<Relationship Id="rId1" Type="${r}/slide" Target="slides/slide1.xml"/>`));
  if (image) {
    zip.file('ppt/media/image1.png', image);
    zip.file('ppt/slides/_rels/slide1.xml.rels', rels(`<Relationship Id="rId2" Type="${r}/image" Target="../media/image1.png"/>`));
  }
  const bg = background ? `<p:bg><p:bgPr>${solid(background)}<a:effectLst/></p:bgPr></p:bg>` : '';
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}"><p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return { name: 'caret.pptx', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function open(page: Page, file: Awaited<ReturnType<typeof deck>>, host = false) {
  await page.goto('/?format=pptx');
  if (host) await page.addStyleTag({ content: `body { color: ${DARK_HOST}; }` });
  await page.locator('input[type=file]').last().setInputFiles(file);
  await inputReady(page);
}

async function caretAt(page: Page, x: number, y: number, key = 'End') {
  const canvas = (await slideCanvas(page).boundingBox())!;
  await page.mouse.click(canvas.x + x * canvas.width / 960, canvas.y + y * canvas.height / 540);
  const input = page.getByRole('textbox', { name: 'Slide text', exact: true });
  await expect(input).toBeFocused();
  await input.press(key);
  await expect(page.getByTestId('pptx-caret')).toBeAttached();
}

async function edit(page: Page, background: string | null, body: string, options: { box?: Box; image?: Buffer; host?: boolean; key?: string } = {}) {
  await open(page, await deck(background, body, options.image), options.host);
  const box = options.box ?? BOX;
  await caretAt(page, box.x + box.w / 2, box.y + box.h / 2, options.key);
}

const caretStyle = (page: Page) => page.getByTestId('pptx-caret').evaluate((element) => {
  const style = getComputedStyle(element);
  return { color: style.backgroundColor, halo: style.boxShadow };
});

async function expectCaret(page: Page, color: string) {
  await expect.poll(async () => (await caretStyle(page)).color).toBe(color);
  expect((await caretStyle(page)).halo).toBe('none');
}

function luminance(color: string) {
  const [red, green, blue] = color.match(/\d+(\.\d+)?/g)!.slice(0, 3).map((value) => {
    const c = Number(value) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * red! + 0.7152 * green! + 0.0722 * blue!;
}

function ratio(first: string, second: string) {
  const [light, dark] = [luminance(first), luminance(second)].sort((x, y) => y - x);
  return (light! + 0.05) / (dark! + 0.05);
}

for (const host of [false, true]) {
  const suffix = host ? ' inside a dark-themed host' : '';
  test(`pptx worker caret is black, not the text colour, for blue text on a white slide${suffix}`, async ({ page }) => {
    await edit(page, null, shape('315EFB'), { host });
    await expectCaret(page, BLACK);
  });

  test(`pptx worker caret is white for white text on a black slide${suffix}`, async ({ page }) => {
    await edit(page, '000000', shape('FFFFFF'), { host });
    await expectCaret(page, WHITE);
  });
}

const fills: [string, string, Parameters<typeof shape>[1], string][] = [
  ['lime text in a navy rectangle', 'C8F56A', { fill: solid('0B1F4B') }, WHITE],
  ['grey text in a navy rectangle', '6B7280', { fill: solid('0B1F4B') }, WHITE],
  ['white text in a mid-grey rectangle', 'FFFFFF', { fill: solid('808080') }, BLACK],
  ['coral text in a peach rounded rectangle', 'E0603A', { fill: solid('FFF0EA'), prst: 'roundRect' }, BLACK],
  ['lime text in a navy rectangle rotated by 15 degrees', 'C8F56A', { fill: solid('0B1F4B'), rot: 15 }, WHITE],
  ['dark text in a near-black rectangle', '111111', { fill: solid('1A1A1A') }, WHITE],
  ['yellow text in a white rectangle with a grey outline', 'FFE600', { fill: solid('FFFFFF'), line: '9CA3AF' }, BLACK],
];

for (const [name, color, options, expected] of fills) {
  test(`pptx worker caret is ${expected === BLACK ? 'black' : 'white'} for ${name}`, async ({ page }) => {
    await edit(page, null, shape(color, options));
    await expectCaret(page, expected);
  });
}

const halves = png(16, (x) => x < 8 ? 0x14 : 0xf0);

test('pptx worker caret is white over the dark half of a picture', async ({ page }) => {
  await edit(page, null, photo + shape('FFFFFF'), { image: halves });
  await expectCaret(page, WHITE);
});

test('pptx worker caret is black over the light half of a picture', async ({ page }) => {
  const box = { x: 576, y: 96, w: 336, h: 96 };
  await edit(page, null, photo + shape('FFFFFF', { box }), { image: halves, box });
  await expectCaret(page, BLACK);
});

test.describe('at a device pixel ratio of 2', () => {
  test.use({ deviceScaleFactor: 2 });

  test('pptx worker caret reads the picture half under it', async ({ page }) => {
    await edit(page, null, photo + shape('FFFFFF'), { image: halves });
    await expectCaret(page, WHITE);
    const box = { x: 576, y: 96, w: 336, h: 96 };
    await edit(page, null, photo + shape('FFFFFF', { box }), { image: halves, box });
    await expectCaret(page, BLACK);
  });
});

const gradient = `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="0B1F4B"/></a:gs><a:gs pos="100000"><a:srgbClr val="F2F2F2"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>`;

test('pptx worker caret is white on the navy end of a navy to near-white gradient', async ({ page }) => {
  await edit(page, null, shape('FFFFFF', { fill: gradient }), { key: 'Home' });
  await expectCaret(page, WHITE);
});

test('pptx worker caret is black on the near-white end of a navy to near-white gradient', async ({ page }) => {
  await edit(page, null, shape('FFFFFF', { fill: gradient, algn: 'r' }));
  await expectCaret(page, BLACK);
});

test('pptx worker caret is black or white against each table cell fill', async ({ page }) => {
  await open(page, await deck(null, table([['0B1F4B', 'C8F56A'], ['FFF0EA', 'E0603A'], ['808080', 'FFFFFF'], ['FFFFFF', '315EFB']])));
  const place = () => page.getByTestId('pptx-caret').evaluate((element) => `${element.style.left} ${element.style.top}`);
  let previous = '';
  for (const [x, y, expected] of [[246, 141, WHITE], [546, 141, BLACK], [246, 231, BLACK], [546, 231, BLACK]] as const) {
    await caretAt(page, x, y);
    await expect.poll(place).not.toBe(previous);
    previous = await place();
    await expectCaret(page, expected);
  }
});

test('pptx worker composition preedit reads against the slide inside a dark-themed host', async ({ page }) => {
  await edit(page, null, shape('172554'), { host: true });
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
