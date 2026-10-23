import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PAGE = { width: 816, height: 1056, margin: 288 };

async function marginDocx(note = false) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>');
  zip.file('word/header1.xml', `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Running head</w:t></w:r></w:p></w:hdr>`);
  zip.file('word/footer1.xml', `<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Running foot</w:t></w:r></w:p></w:ftr>`);
  zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1"><w:p><w:r><w:t>Bottom note</w:t></w:r></w:p></w:footnote></w:footnotes>`);
  const reference = note ? '<w:r><w:footnoteReference w:id="1"/></w:r>' : '';
  const paragraphs = ['First body', 'Middle body', 'Last body'].map((text, index) => `<w:p><w:r><w:t>${text}</w:t></w:r>${index === 2 ? reference : ''}</w:p>`).join('');
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraphs}<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="4320" w:bottom="4320" w:left="1440" w:right="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, note = false, bottom = false) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'margin-click.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await marginDocx(note) });
  const canvas = page.locator('.paged-editor__page canvas').first();
  await expect(canvas).toBeVisible({ timeout: 60_000 });
  if (bottom) await canvas.evaluate((node) => node.scrollIntoView({ block: 'end' }));
  const box = (await canvas.boundingBox())!;
  expect(Math.round(box.width)).toBe(PAGE.width);
  return box;
}

async function saved(page: Page, info: TestInfo) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath('margin-click.docx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const text = async (part: string) => (await zip.file(part)!.async('string'));
  const body = await text('word/document.xml');
  const paragraphs = body.split('</w:p>').slice(0, -1).map((xml) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join(''));
  return { paragraphs, header: await text('word/header1.xml'), footer: await text('word/footer1.xml'), notes: await text('word/footnotes.xml') };
}

const focusedInput = (page: Page) => page.evaluate(() => document.activeElement?.classList.contains('paged-editor__peerless-input') ?? false);

test('single click in the header band focuses the input and types at the start of the first body paragraph', async ({ page }, info) => {
  const box = await open(page);
  await page.mouse.click(box.x + 40, box.y + 160);
  expect(await focusedInput(page)).toBe(true);
  await page.keyboard.type('Qm1 ');
  const { paragraphs, header } = await saved(page, info);
  expect(paragraphs[0]).toBe('Qm1 First body');
  expect(paragraphs.slice(1)).toEqual(['Middle body', 'Last body']);
  expect(header).not.toContain('Qm1');
});

test('single click in the footer band focuses the input and types at the start of the last body paragraph', async ({ page }, info) => {
  const box = await open(page, false, true);
  expect(box.y + PAGE.height - 160).toBeLessThan(page.viewportSize()!.height);
  await page.mouse.click(box.x + 40, box.y + PAGE.height - 160);
  expect(await focusedInput(page)).toBe(true);
  await page.keyboard.type('Qm2 ');
  const { paragraphs, footer } = await saved(page, info);
  expect(paragraphs).toEqual(['First body', 'Middle body', 'Qm2 Last body']);
  expect(footer).not.toContain('Qm2');
});

test('single click in the footer band above a bottom footnote area types into the body, not the note', async ({ page }, info) => {
  const box = await open(page, true, true);
  expect(box.y + PAGE.height - 160).toBeLessThan(page.viewportSize()!.height);
  await page.mouse.click(box.x + 40, box.y + PAGE.height - 160);
  expect(await focusedInput(page)).toBe(true);
  await page.keyboard.type('Qm4 ');
  const { paragraphs, notes } = await saved(page, info);
  expect(paragraphs.slice(0, 3)).toEqual(['First body', 'Middle body', 'Qm4 Last body']);
  expect(notes).not.toContain('Qm4');
});

test('double click in the header band still opens the header', async ({ page }, info) => {
  const box = await open(page);
  await page.mouse.dblclick(box.x + 200, box.y + 60);
  const chrome = page.locator('.hf-inline-editor');
  await expect(chrome).toContainText('Header');
  await page.keyboard.press('ControlOrMeta+Home');
  await page.keyboard.type('Qm3 ');
  const { paragraphs, header } = await saved(page, info);
  expect(header).toContain('Qm3 Running head');
  expect(paragraphs).toEqual(['First body', 'Middle body', 'Last body']);
});
