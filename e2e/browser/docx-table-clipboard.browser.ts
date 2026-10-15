import { expect, test, type Locator, type Page, type TestInfo } from 'playwright/test';
import { pinnedEvaluate } from './pinned-evaluate';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

const LABELS = ['RIGHT', 'INDEX', 'RATIO', 'VPS'];
const ARMS = [{ name: 'default', query: '' }, { name: 'worker', query: '&peerless=1' }] as const;
type Arm = (typeof ARMS)[number];

function documentXml(rows: string[][]): string {
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr><w:p>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p></w:tc>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid>${'<w:gridCol w:w="2340"/>'.repeat(4)}</w:tblGrid>${rows.map(row => `<w:tr>${row.map(cell).join('')}</w:tr>`).join('')}</w:tbl>`;
  return '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    `<w:p><w:r><w:t>Before</w:t></w:r></w:p>${table}<w:p/>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>';
}

async function open(page: Page, arm: Arm, rows: string[][]): Promise<Locator> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', documentXml(rows));
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto('/?format=docx' + arm.query);
  await page.locator('input[type=file]').last().setInputFiles({ name: 'table-clipboard.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.canvas-page canvas, .paged-editor__page canvas').first()).toBeVisible({ timeout: 60_000 });
  const input = page.locator('[data-testid=yrs-input], .paged-editor__peerless-input');
  await input.focus();
  await expect(input).toBeEditable();
  return input;
}

async function savedTables(page: Page, info: TestInfo, name: string): Promise<string[][][]> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`${name}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate((xml) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const body = document.getElementsByTagNameNS('*', 'body')[0]!;
    const children = (node: Element, name: string) => Array.from(node.children).filter(child => child.localName === name);
    return children(body, 'tbl').map(table => children(table, 'tr').map(row => children(row, 'tc').map(cell =>
      Array.from(cell.getElementsByTagNameNS('*', 't')).map(text => text.textContent ?? '').join(''))));
  }, xml);
}

async function copiedFlavors(page: Page): Promise<{ types: string[]; html: string; text: string }> {
  let flavors = { types: [] as string[], html: '', text: '' };
  await expect.poll(async () => {
    flavors = await pinnedEvaluate(page, async () => {
      const items = await navigator.clipboard.read();
      const types = items.flatMap(item => [...item.types]);
      const read = async (type: string) => {
        const item = items.find(entry => entry.types.includes(type));
        return item ? (await item.getType(type)).text() : '';
      };
      return { types, html: await read('text/html'), text: await read('text/plain') };
    });
    return flavors.text.includes('RIGHT');
  }).toBe(true);
  return flavors;
}

async function selectTable(page: Page, input: Locator, arm: Arm) {
  if (arm.name === 'worker') {
    await input.press('ControlOrMeta+a');
    return;
  }
  await input.press('ControlOrMeta+Home');
  await input.press('ArrowDown');
  await page.getByRole('button', { name: 'More table options', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Select entire table' }).click();
  await input.focus();
}

async function caretAfterTable(input: Locator, arm: Arm) {
  if (arm.name === 'default') await input.press('ArrowUp');
  await input.press('ControlOrMeta+End');
}

async function copyPasteAfterTable(page: Page, input: Locator, arm: Arm) {
  await pinnedEvaluate(page, () => navigator.clipboard.writeText('sentinel'));
  await selectTable(page, input, arm);
  await input.press('ControlOrMeta+c');
  const flavors = await copiedFlavors(page);
  expect(flavors.types).toContain('text/html');
  expect(flavors.html).toContain('<table');
  await caretAfterTable(input, arm);
  await input.press('ControlOrMeta+v');
}

for (const arm of ARMS) {
  test(`${arm.name}: a copied table pastes after itself as a second table in one undo step`, async ({ page }, info) => {
    const rows = [LABELS, ...Array.from({ length: 4 }, () => ['', '', '', ''])];
    const input = await open(page, arm, rows);
    await copyPasteAfterTable(page, input, arm);
    await expect.poll(async () => (await savedTables(page, info, `${arm.name}-pasted`)).length).toBe(2);
    const tables = await savedTables(page, info, `${arm.name}-pasted-final`);
    expect(tables[0]).toEqual(rows);
    expect(tables[1]).toEqual(rows);
    await input.focus();
    await input.press('ControlOrMeta+z');
    await expect.poll(async () => (await savedTables(page, info, `${arm.name}-undone`)).length).toBe(1);
  });

  test(`${arm.name}: a table spanning pages round-trips through the clipboard`, async ({ page }, info) => {
    const rows = [LABELS, ...Array.from({ length: 69 }, (_, row) => LABELS.map((_, column) => `r${row + 1}c${column + 1}`))];
    const input = await open(page, arm, rows);
    await copyPasteAfterTable(page, input, arm);
    await expect.poll(async () => (await savedTables(page, info, `${arm.name}-long`)).length, { timeout: 30_000 }).toBe(2);
    const tables = await savedTables(page, info, `${arm.name}-long-final`);
    expect(tables[0]).toEqual(rows);
    expect(tables[1]).toEqual(rows);
  });
}
