import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const DARK_HOST = 'rgb(237, 237, 237)';
const PLAIN = '<w:p><w:r><w:t>Caret on a white page</w:t></w:r></w:p>';
const cell = (shading: string, run: string) => '<w:tbl><w:tblPr><w:tblW w:w="5760" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5760"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5760" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="' + shading + '"/></w:tcPr><w:p>' + run + '</w:p></w:tc></w:tr></w:tbl><w:p/>';
const DARK_CELL = cell('1F1F1F', '<w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>Dark</w:t></w:r>');
const HIGHLIGHT = '<w:p><w:r><w:rPr><w:color w:val="FFFFFF"/><w:highlight w:val="darkBlue"/></w:rPr><w:t>Highlighted</w:t></w:r></w:p>';

async function open(page: Page, arm: string, body: string, host?: string, back = 0) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + body + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>');
  await page.goto('/?format=docx' + arm);
  if (host) await page.addStyleTag({ content: `body { color: ${host}; }` });
  await page.locator('input[type=file]').last().setInputFiles({ name: 'caret.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  const surface = page.locator('.canvas-page, .paged-editor__page').first();
  await expect(surface.locator('canvas').first()).toBeVisible({ timeout: 60_000 });
  if (host) await page.evaluate((color) => { (document.querySelector('.oox-root')!.parentElement as HTMLElement).style.color = color; }, host);
  await surface.click({ position: { x: 140, y: 100 } });
  const input = page.locator('[data-testid=yrs-input], .paged-editor__input').first();
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  await input.press('End');
  for (let step = 0; step < back; step += 1) await input.press('ArrowLeft');
}

async function caretBox(page: Page) {
  const caret = page.locator('[data-testid=caret], [data-docx-pages-host] > div[aria-hidden][style*="will-change: transform"]').first();
  await expect(caret).toBeAttached();
  await expect.poll(async () => (await caret.boundingBox())?.height ?? 0).toBeGreaterThan(4);
  return (await caret.boundingBox())!;
}

async function pixels(page: Page, clip: { x: number; y: number; width: number; height: number }) {
  const png = (await page.screenshot({ clip })).toString('base64');
  return page.evaluate(async (data) => {
    const image = new Image();
    image.src = 'data:image/png;base64,' + data;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = image.width;
    canvas.height = image.height;
    const context = canvas.getContext('2d')!;
    context.drawImage(image, 0, 0);
    const { data: rgba } = context.getImageData(0, 0, image.width, image.height);
    const out: [number, number, number][] = [];
    for (let index = 0; index < rgba.length; index += 4) out.push([rgba[index]!, rgba[index + 1]!, rgba[index + 2]!]);
    return out;
  }, png);
}

function luminance([r, g, b]: [number, number, number]) {
  const channel = (value: number) => {
    const c = value / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
}

function contrast(a: [number, number, number], b: [number, number, number]) {
  const [light, dark] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (light! + 0.05) / (dark! + 0.05);
}

async function caretContrast(page: Page) {
  const box = await caretBox(page);
  const clip = { x: Math.floor(box.x) - 1, y: Math.round(box.y + box.height / 2) - 1, width: Math.ceil(box.width) + 2, height: 2 };
  const frames: [number, number, number][][] = [];
  let best = 0;
  for (let attempt = 0; attempt < 16 && best < 3; attempt += 1) {
    const frame = await pixels(page, clip);
    for (const seen of frames) frame.forEach((pixel, index) => { best = Math.max(best, contrast(pixel, seen[index]!)); });
    frames.push(frame);
    await page.waitForTimeout(140);
  }
  return best;
}

for (const [arm, label] of [['', 'worker'], ['&peerless=0', 'non-worker']] as const) {
  test(`docx caret stays visible on the page inside a dark-themed host (${label})`, async ({ page }) => {
    await open(page, arm, PLAIN, DARK_HOST);
    expect(await caretContrast(page)).toBeGreaterThanOrEqual(3);
  });
}

test('docx worker caret follows white text in a dark-shaded table cell', async ({ page }) => {
  await open(page, '', DARK_CELL);
  expect(await caretContrast(page)).toBeGreaterThanOrEqual(3);
});

test('docx worker caret follows white text in a dark-shaded cell inside a dark-themed host', async ({ page }) => {
  await open(page, '', DARK_CELL, DARK_HOST);
  expect(await caretContrast(page)).toBeGreaterThanOrEqual(3);
});

test('docx worker caret follows white text on a dark highlight', async ({ page }) => {
  await open(page, '', HIGHLIGHT, undefined, 4);
  expect(await caretContrast(page)).toBeGreaterThanOrEqual(3);
});
