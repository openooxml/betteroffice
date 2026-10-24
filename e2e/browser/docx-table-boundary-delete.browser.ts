import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { pagedDocx } from '../../packages/docx-react/src/components/DocxEditor/__fixtures__/pagedDocx';
import { renderedBox } from './rendered-box';

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const root = '[data-docx-worker-editor]';
const input = (page: Page) => page.locator(`${root} .paged-editor__input`);
const mirror = (page: Page) => page.locator(`${root} .canvas-page-mirror`).first();
const firstTable = [['T1-A1', 'T1-B1'], ['T1-A2', 'T1-B2']];
const secondTable = [['T2-A1', 'T2-B1'], ['T2-A2', 'T2-B2']];

type BodyBlock = { kind: 'p'; text: string } | { kind: 'tbl'; cells: string[][] };

function paragraph(text: string, id?: string): string {
  return `<w:p${id ? ` w14:paraId="${id}"` : ''}><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
    (text ? '<w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr>' +
      `<w:t>${text}</w:t></w:r>` : '') + '</w:p>';
}

function table(cells: string[][]): string {
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side =>
    `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`
  ).join('');
  return '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>' +
    `<w:tblBorders>${borders}</w:tblBorders></w:tblPr>` +
    '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' +
    cells.map(row => '<w:tr>' + row.map(text => '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>' +
      paragraph(text) + '</w:tc>').join('') + '</w:tr>').join('') + '</w:tbl>';
}

async function fixture(twoTables: boolean): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await pagedDocx(1));
  const before = twoTables ? 'Paragraph before the tables.' : 'Paragraph before the table.';
  const after = twoTables ? 'Paragraph after the tables.' : 'Paragraph after the table.';
  zip.file('word/document.xml', `<w:document xmlns:w="${wordNamespace}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>` +
    paragraph(before, '00000001') + table(firstTable) + paragraph(twoTables ? '' : after, '00000003') +
    (twoTables ? table(secondTable) + paragraph(after, '00000005') : '') +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>' +
    '</w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function open(page: Page, buffer: Buffer) {
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'table-boundary.docx', mimeType, buffer });
  await expect(page.locator(`${root} .paged-editor__page canvas`).first()).toBeVisible({ timeout: 60_000 });
  await expect(input(page)).toBeEditable();
  await expect(mirror(page)).toContainText('T1-B2');
  await expect(page.locator('.notice')).toHaveCount(0);
}

async function targetParagraphStart(page: Page) {
  const run = mirror(page).locator('.layout-paragraph[data-para-id="00000003"] .layout-run-text').first();
  await expect(run).toBeAttached();
  const box = await renderedBox(run);
  await page.mouse.click(box.x + Math.min(box.width / 2, 2), box.y + box.height / 2);
  await input(page).press('Home');
  const caret = page.locator(`${root} div[aria-hidden="true"][style*="currentcolor" i]`);
  await expect(caret).toHaveCount(1);
  await expect.poll(async () => {
    const bounds = await caret.boundingBox();
    return bounds ? Math.abs(bounds.y + bounds.height / 2 - box.y - box.height / 2) : Infinity;
  }).toBeLessThan(box.height);
}

async function save(page: Page): Promise<{ buffer: Buffer; blocks: BodyBlock[] }> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const stream = await (await downloading).createReadStream();
  if (!stream) throw new Error('Save did not produce a download');
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  const buffer = Buffer.concat(chunks);
  const zip = await JSZip.loadAsync(buffer);
  const part = zip.file('word/document.xml');
  if (!part) throw new Error('Saved document has no word/document.xml');
  const xml = await part.async('string');
  const blocks = await page.evaluate(({ xml, wordNamespace }): BodyBlock[] => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Invalid saved document.xml');
    const body = document.getElementsByTagNameNS(wordNamespace, 'body')[0];
    if (!body) throw new Error('Saved document has no body');
    const children = (node: Element, name: string) => Array.from(node.children).filter(child => child.localName === name);
    const text = (node: Element) => Array.from(node.getElementsByTagNameNS(wordNamespace, 't'))
      .map(run => run.textContent ?? '').join('');
    return Array.from(body.children).filter(node => node.localName !== 'sectPr').map((node): BodyBlock => {
      if (node.localName === 'p') {
        if (node.getElementsByTagNameNS(wordNamespace, 'tbl').length) throw new Error('Table nested inside a paragraph');
        return { kind: 'p', text: text(node) };
      }
      if (node.localName !== 'tbl') throw new Error(`Unexpected body block ${node.localName}`);
      return {
        kind: 'tbl',
        cells: children(node, 'tr').map(row => children(row, 'tc').map(cell => {
          if (!children(cell, 'p').length) throw new Error('Table cell has no paragraph');
          return text(cell);
        })),
      };
    });
  }, { xml, wordNamespace });
  return { buffer, blocks };
}

function expectedBlocks(twoTables: boolean, typingInCell: boolean): BodyBlock[] {
  const cells = firstTable.map(row => [...row]);
  if (typingInCell) cells[1]![1] += 'Hello';
  const before = twoTables ? 'Paragraph before the tables.' : 'Paragraph before the table.';
  const after = twoTables ? 'Paragraph after the tables.' : 'Paragraph after the table.';
  return [
    { kind: 'p', text: before },
    { kind: 'tbl', cells },
    { kind: 'p', text: (typingInCell ? '' : 'Hello') + (twoTables ? '' : after) },
    ...(twoTables ? [{ kind: 'tbl' as const, cells: secondTable }, { kind: 'p' as const, text: after }] : []),
  ];
}

for (const twoTables of [false, true]) {
  test(`default worker Backspace ${twoTables ? 'between tables preserves both tables' : 'after a table preserves block structure'}`, async ({ page }) => {
    await open(page, await fixture(twoTables));
    await targetParagraphStart(page);
    await input(page).press('Backspace');
    await page.keyboard.insertText('Hello');
    await expect(mirror(page)).toContainText('Hello');
    for (const text of [...firstTable.flat(), ...(twoTables ? secondTable.flat() : [])]) {
      await expect(mirror(page)).toContainText(text);
    }
    await expect(page.locator('.notice')).toHaveCount(0);
    const saved = await save(page);
    if (twoTables) expect(saved.blocks).toEqual(expectedBlocks(true, true));
    else expect([expectedBlocks(false, false), expectedBlocks(false, true)]).toContainEqual(saved.blocks);
    await open(page, saved.buffer);
    await expect(mirror(page)).toContainText('Hello');
    expect((await save(page)).blocks).toEqual(saved.blocks);
    await expect(page.locator('.notice')).toHaveCount(0);
  });
}
