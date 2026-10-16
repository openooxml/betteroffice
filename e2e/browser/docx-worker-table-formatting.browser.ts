import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

function documentXml(): string {
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const rows = [0, 1].map((row) => `<w:tr>${[0, 1].map((column) => cell(`R${row}C${column}`)).join('')}</w:tr>`).join('');
  return '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    `<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid>${'<w:gridCol w:w="2340"/>'.repeat(2)}</w:tblGrid>${rows}</w:tbl><w:p/>` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>';
}

async function open(page: Page) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', documentXml());
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'table-formatting.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  await expect(page.locator('.paged-editor__page canvas').first()).toBeVisible({ timeout: 60000 });
  const input = page.locator('.paged-editor__peerless-input');
  await input.focus();
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  await input.press('ArrowDown');
  return input;
}

async function tableAction(page: Page, action: unknown) {
  const result = await page.evaluate((value) => {
    const harness = (window as unknown as { __docxPeerlessHarness?: { tableAction(a: unknown): Promise<unknown> } }).__docxPeerlessHarness;
    if (!harness) throw new Error('DOCX harness is unavailable');
    return harness.tableAction(value);
  }, action);
  expect(result).toMatchObject({ ok: true });
}

async function savedTable(page: Page, info: TestInfo, step: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(`table-formatting-${step}.docx`);
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate((source) => {
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const doc = new DOMParser().parseFromString(source, 'application/xml');
    const table = doc.getElementsByTagNameNS('*', 'tbl')[0]!;
    const cells = Array.from(table.getElementsByTagNameNS('*', 'tc')).map((cell) => {
      const borders = cell.getElementsByTagNameNS('*', 'tcBorders')[0];
      const sides: Record<string, string> = {};
      if (borders) for (const side of Array.from(borders.children)) {
        sides[side.localName] = `${side.getAttributeNS(w, 'val')}/${side.getAttributeNS(w, 'sz')}/${side.getAttributeNS(w, 'color')}`;
      }
      return sides;
    });
    const tblBorders = table.getElementsByTagNameNS('*', 'tblBorders')[0];
    const tblW = table.getElementsByTagNameNS('*', 'tblW')[0];
    const jc = table.getElementsByTagNameNS('*', 'jc')[0];
    return {
      cells,
      tblBorders: tblBorders ? tblBorders.children.length : 0,
      tblStyle: table.getElementsByTagNameNS('*', 'tblStyle').length,
      tblW: tblW ? `${tblW.getAttributeNS(w, 'w')}/${tblW.getAttributeNS(w, 'type')}` : null,
      jc: jc?.getAttributeNS(w, 'val') ?? null,
    };
  }, xml);
}

test('worker borders preserve sides and saved shared edges', async ({ page }, info) => {
  await open(page);
  await tableAction(page, 'borderNone');
  let saved = await savedTable(page, info, 'none');
  for (const cell of saved.cells) {
    expect(cell.top).toBe('none/0/000000');
    expect(cell.bottom).toBe('none/0/000000');
    expect(cell.left).toBe('none/0/000000');
    expect(cell.right).toBe('none/0/000000');
  }
  expect(saved.tblBorders).toBe(6);
  expect(saved.tblStyle).toBe(0);

  await tableAction(page, { type: 'borderColor', color: 'FF0000' });
  saved = await savedTable(page, info, 'color');
  expect(saved.cells[0]!.top).toBe('single/4/FF0000');

  await tableAction(page, { type: 'borderWidth', size: 12 });
  saved = await savedTable(page, info, 'width');
  expect(saved.cells[0]!.top).toBe('single/12/FF0000');

  await tableAction(page, 'borderOutside');
  saved = await savedTable(page, info, 'outside');
  expect(saved.cells[0]!.top).toBe('single/12/FF0000');
  expect(saved.cells[0]!.left).toBe('single/12/FF0000');

  await tableAction(page, { type: 'cellBorder', side: 'top', style: 'double', size: 6, color: '00FF00' });
  saved = await savedTable(page, info, 'side');
  expect(saved.cells[0]!.top).toBe('double/6/00FF00');
  expect(saved.cells[1]!.top).toBe('double/6/00FF00');

  await tableAction(page, { type: 'applyTableStyle', styleId: 'TableGrid' });
  saved = await savedTable(page, info, 'style');
  expect(saved.cells[0]!.top).toBe('single/4/000000');
  expect(saved.tblStyle).toBe(0);

  await tableAction(page, { type: 'applyTableStyle', styleId: 'TableNormal' });
  saved = await savedTable(page, info, 'nostyle');
  expect(saved.cells[0]!.top).toBe('single/4/000000');
  expect(saved.tblStyle).toBe(0);
});

test('worker table properties dialog reads and applies after typing', async ({ page }, info) => {
  const input = await open(page);
  await input.press('x');
  await tableAction(page, { type: 'openTableProperties' });
  const dialog = page.getByRole('dialog', { name: 'Table Properties', exact: true });
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('select').first()).toHaveValue('dxa');
  await dialog.locator('select').first().selectOption('pct');
  await dialog.locator('input[type=number]').fill('2500');
  await dialog.locator('select').nth(1).selectOption('center');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(dialog).toBeHidden();
  let saved = await savedTable(page, info, 'props');
  expect(saved.tblW).toBe('2500/pct');
  expect(saved.jc).toBe('center');

  await tableAction(page, { type: 'openTableProperties' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(dialog).toBeHidden();
  saved = await savedTable(page, info, 'cancel');
  expect(saved.tblW).toBe('2500/pct');
  expect(saved.jc).toBe('center');
});
