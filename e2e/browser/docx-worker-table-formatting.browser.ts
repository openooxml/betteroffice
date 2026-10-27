import { expect, test, type Page, type TestInfo } from 'playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

function documentXml(): string {
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`;
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const rows = [0, 1].map((row) => `<w:tr>${[0, 1].map((column) => cell(`R${row}C${column}`)).join('')}</w:tr>`).join('');
  return '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    `<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="4680" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders></w:tblPr>` +
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
  const input = page.locator('.paged-editor__input');
  await input.focus();
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  await input.press('ArrowDown');
  await expect(page.getByTestId('toolbar-table-borders')).toBeEnabled();
  return input;
}

async function expectHighlights(page: Page, count: number) {
  await page.waitForFunction((expected) => [...document.querySelectorAll('div[aria-hidden="true"]')]
    .filter((element) => (element as HTMLElement).style.background.includes('59, 130, 246')).length === expected,
    count, { timeout: 30000 });
}

async function selectEntireTable(page: Page, input: ReturnType<Page['locator']>) {
  await input.focus();
  await input.press('ControlOrMeta+Home');
  for (let attempt = 0; attempt < 4; attempt += 1) {
    await input.press('ArrowDown');
    await page.locator('.paged-editor__page').first()
      .click({ button: 'right', position: { x: 100, y: 100 }, force: true });
    const item = page.getByRole('menuitem', { name: 'Select entire table', exact: true });
    if (await item.count() > 0) {
      await item.click();
      await expectHighlights(page, 4);
      return;
    }
    await page.keyboard.press('Escape');
  }
  throw new Error('Caret did not enter the table');
}

async function overflowItem(page: Page, path: string[]) {
  await page.getByTestId('toolbar-more').click();
  let menu = page.getByRole('menu', { name: 'More actions', exact: true });
  await expect(menu).toBeVisible();
  for (const label of path.slice(0, -1)) {
    await menu.locator(`[data-label="${label}"]`).focus();
    await page.keyboard.press('ArrowRight');
    menu = page.getByRole('menu', { name: label, exact: true });
    await expect(menu).toBeVisible();
  }
  await menu.locator(`[data-label="${path.at(-1)}"]`).click();
}

async function borderPreset(page: Page, label: string) {
  const picker = page.getByTestId('toolbar-table-borders');
  if (await picker.isVisible()) {
    await picker.click();
    await page.locator(`button[title="${label}"]`).click();
    return;
  }
  await overflowItem(page, ['Table', 'Borders', label]);
}

async function borderColorRed(page: Page) {
  const arrow = page.locator('.docx-color-picker-arrow[aria-label="Border Color"]');
  if (await arrow.isVisible()) {
    await arrow.click();
    const picker = page.getByRole('dialog', { name: 'Border Color picker', exact: true });
    await expect(picker).toBeVisible();
    await picker.getByLabel('Custom hex color', { exact: true }).fill('FF0000');
    await picker.getByRole('button', { name: 'Apply', exact: true }).click();
    return;
  }
  await overflowItem(page, ['Table', 'Border Color', 'Red']);
}

async function borderWidth(page: Page, label: string) {
  const picker = page.getByTestId('toolbar-table-border-width');
  if (await picker.isVisible()) {
    await picker.click();
    await page.getByRole('button', { name: label, exact: true }).click();
    return;
  }
  await overflowItem(page, ['Table', 'Border width', label]);
}

async function openTableProperties(page: Page) {
  const more = page.getByTestId('toolbar-table-more');
  if (await more.isVisible()) {
    await more.click();
    await page.getByRole('menuitem', { name: 'Table properties...', exact: true }).click();
    return;
  }
  await overflowItem(page, ['Table', 'Table properties...']);
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
    const tblStyle = table.getElementsByTagNameNS('*', 'tblStyle')[0];
    const tblW = table.getElementsByTagNameNS('*', 'tblW')[0];
    const jc = table.getElementsByTagNameNS('*', 'jc')[0];
    return {
      cells,
      tblBorders: tblBorders ? tblBorders.children.length : 0,
      tblStyle: tblStyle?.getAttributeNS(w, 'val') ?? null,
      tblW: tblW ? `${tblW.getAttributeNS(w, 'w')}/${tblW.getAttributeNS(w, 'type')}` : null,
      jc: jc?.getAttributeNS(w, 'val') ?? null,
    };
  }, xml);
}

test('worker borders preserve sides and saved shared edges', async ({ page }, info) => {
  const input = await open(page);
  await borderPreset(page, 'No borders');
  let saved = await savedTable(page, info, 'single-none');
  expect(saved.cells[0]).toEqual({ top: 'none/null/null', bottom: 'none/null/null', left: 'none/null/null', right: 'none/null/null' });
  const line = 'single/4/000000';
  expect(saved.cells[1]).toEqual({ top: line, bottom: line, left: 'none/null/null', right: line });
  expect(saved.cells[2]).toEqual({ top: 'none/null/null', bottom: line, left: line, right: line });
  expect(saved.cells[3]).toEqual({});
  expect(saved.tblBorders).toBe(6);
  expect(saved.tblStyle).toBe('TableGrid');

  await selectEntireTable(page, input);
  await borderPreset(page, 'No borders');
  saved = await savedTable(page, info, 'none');
  for (const cell of saved.cells) {
    expect(cell.top).toBe('none/null/null');
    expect(cell.bottom).toBe('none/null/null');
    expect(cell.left).toBe('none/null/null');
    expect(cell.right).toBe('none/null/null');
  }
  expect(saved.tblBorders).toBe(6);
  expect(saved.tblStyle).toBe('TableGrid');

  await borderColorRed(page);
  saved = await savedTable(page, info, 'color');
  expect(saved.cells[0]!.top).toBe('single/4/FF0000');
  expect(saved.cells[3]!.bottom).toBe('single/4/FF0000');

  await borderWidth(page, '1.5 pt');
  saved = await savedTable(page, info, 'width');
  expect(saved.cells[0]!.top).toBe('single/12/FF0000');

  await borderPreset(page, 'Outside borders');
  saved = await savedTable(page, info, 'outside');
  expect(saved.cells[0]!.top).toBe('single/12/FF0000');
  expect(saved.cells[0]!.left).toBe('single/12/FF0000');
  expect(saved.tblStyle).toBe('TableGrid');
});

test('worker table properties dialog reads and applies after typing', async ({ page }, info) => {
  const input = await open(page);
  await input.press('x');
  await openTableProperties(page);
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

  await input.focus();
  await openTableProperties(page);
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(dialog).toBeHidden();
  saved = await savedTable(page, info, 'cancel');
  expect(saved.tblW).toBe('2500/pct');
  expect(saved.jc).toBe('center');
});
