import { test, expect, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const seed = 'original';

type Cell = { type: 's' | 'n' | 'b' | 'str'; value: string; formula: string | null };
type Entry = {
  address: string;
  input: string;
  cell: Cell;
  display?: string | RegExp;
  source?: 'cell' | 'formula';
};

const textCell = (value: string): Cell => ({ type: 's', value, formula: null });
const numberCell = (value: string): Cell => ({ type: 'n', value, formula: null });

async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${relationships}"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Input" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${relationships}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${relationships}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`);
  zip.file('xl/sharedStrings.xml', `<sst xmlns="${main}" count="4" uniqueCount="1"><si><t>${seed}</t></si></sst>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><dimension ref="A1:D2"/><sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/><cols><col min="1" max="4" width="28" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>0</v></c><c r="D1" t="s"><v>0</v></c></row><row r="2"><c r="D2"><v>0</v></c></row></sheetData></worksheet>`);
  return { name: 'input.xlsx', mimeType, buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function open(page: Page, file?: string) {
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles(file ?? await fixture());
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(page.getByTestId('xlsx-formula-input')).toBeEditable();
  await expect(page.locator('.notice')).toHaveCount(0);
  await expect(page.getByTestId('xlsx-error')).toHaveCount(0);
}

async function select(page: Page, address: string) {
  const match = /^([A-D])([12])$/.exec(address);
  expect(match).not.toBeNull();
  const scroll = page.getByTestId('xlsx-scroll');
  await scroll.press('ControlOrMeta+Home');
  for (let col = 0; col < match![1].charCodeAt(0) - 65; col++) await scroll.press('ArrowRight');
  if (match![2] === '2') await scroll.press('ArrowDown');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue(address);
}

async function startEdit(page: Page, source: 'cell' | 'formula' = 'cell') {
  if (source === 'cell') await page.getByTestId('xlsx-scroll').press('F2');
  const input = page.getByTestId(source === 'cell' ? 'xlsx-cell-editor' : 'xlsx-formula-input');
  if (source === 'formula') await input.click();
  else await input.focus();
  await expect(input).toHaveValue(seed);
  await input.press('ControlOrMeta+A');
  return input;
}

async function visible(page: Page, address: string, value: string | RegExp) {
  await select(page, address);
  const body = typeof value === 'string'
    ? value.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+')
    : value.source;
  const text = new RegExp(`^${address},\\s*(?:${body})\\s*(?:,\\s*selected)?$`);
  await expect(page.getByRole('gridcell').filter({ hasText: text })).toHaveCount(1);
}

async function expectEntry(page: Page, entry: Entry) {
  await visible(page, entry.address, entry.display ?? entry.cell.value);
  const input = entry.cell.formula !== null ? '=' + entry.cell.formula :
    entry.cell.type === 'b' ? entry.cell.value === '1' ? 'TRUE' : 'FALSE' :
      entry.input.startsWith("'") ? entry.input : entry.cell.value;
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(input);
}

async function save(page: Page, info: TestInfo, label: string) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const file = info.outputPath(label + '.xlsx');
  await (await downloading).saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const strings = await zip.file('xl/sharedStrings.xml')?.async('string') ?? '';
  const cells = await page.evaluate(({ sheet, strings }) => {
    const parse = (xml: string) => {
      const document = new DOMParser().parseFromString(xml, 'application/xml');
      if (document.querySelector('parsererror')) throw new Error('Invalid saved XML');
      return document;
    };
    const text = (node: Element) => Array.from(node.getElementsByTagNameNS('*', 't'))
      .map((part) => part.textContent ?? '').join('');
    const shared = strings ? Array.from(parse(strings).getElementsByTagNameNS('*', 'si')).map(text) : [];
    return Object.fromEntries(Array.from(parse(sheet).getElementsByTagNameNS('*', 'c')).map((cell) => {
      const type = cell.getAttribute('t') ?? 'n';
      const raw = cell.getElementsByTagNameNS('*', 'v')[0]?.textContent ?? '';
      if (type === 's' && (!/^\d+$/.test(raw) || shared[Number(raw)] === undefined)) {
        throw new Error('Invalid shared string index');
      }
      return [cell.getAttribute('r')!, {
        type: type === 'inlineStr' ? 's' : type,
        value: type === 's' ? shared[Number(raw)] : type === 'inlineStr' ? text(cell) : raw,
        formula: cell.getElementsByTagNameNS('*', 'f')[0]?.textContent ?? null,
      }];
    }));
  }, { sheet, strings });
  return { file, cells };
}

async function history(page: Page, info: TestInfo, entry: Entry) {
  await expectEntry(page, entry);
  expect((await save(page, info, entry.address + '-edited')).cells[entry.address]).toEqual(entry.cell);
  const undo = page.getByTestId('xlsx-undo');
  await expect(undo).toBeEnabled();
  if (entry.source === 'formula') await page.getByTestId('xlsx-scroll').press('ControlOrMeta+z');
  else await undo.click();
  await visible(page, entry.address, seed);
  if (entry.address === 'A1') await expect(undo).toBeDisabled();
  expect((await save(page, info, entry.address + '-undone')).cells[entry.address]).toEqual(textCell(seed));
  const redo = page.getByTestId('xlsx-redo');
  await expect(redo).toBeEnabled();
  if (entry.source === 'formula') await page.getByTestId('xlsx-scroll').press('ControlOrMeta+Shift+z');
  else await redo.click();
  await expectEntry(page, entry);
  expect((await save(page, info, entry.address + '-redone')).cells[entry.address]).toEqual(entry.cell);
}

async function reopen(page: Page, info: TestInfo, entries: Entry[]) {
  const saved = await save(page, info, 'before-reopen');
  for (const entry of entries) expect(saved.cells[entry.address]).toEqual(entry.cell);
  await open(page, saved.file);
  for (const entry of entries) await expectEntry(page, entry);
  const reopened = await save(page, info, 'after-reopen');
  for (const entry of entries) expect(reopened.cells[entry.address]).toEqual(entry.cell);
}

for (const source of ['cell', 'formula'] as const) {
  test(`xlsx input: ${source} CJK composition commits once with Enter`, async ({ page }, info) => {
    await open(page);
    const input = await startEdit(page, source);
    const cdp = await page.context().newCDPSession(page);
    await cdp.send('Input.imeSetComposition', { text: 'かんじ', selectionStart: 3, selectionEnd: 3 });
    await expect(input).toHaveValue('かんじ');
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await cdp.send('Input.imeSetComposition', { text: '漢字かな', selectionStart: 4, selectionEnd: 4 });
    await expect(input).toHaveValue('漢字かな');
    await cdp.send('Input.insertText', { text: '漢字かな' });
    await expect(input).toHaveValue('漢字かな');
    await input.press('Enter');
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A2');
    const entry: Entry = { address: 'A1', input: '漢字かな', cell: textCell('漢字かな'), source };
    await history(page, info, entry);
    await reopen(page, info, [entry]);
    await cdp.detach();
  });

  test(`xlsx input: ${source} CJK composition cancels with Escape`, async ({ page }, info) => {
    await open(page);
    const input = await startEdit(page, source);
    await page.keyboard.insertText('uncommitted');
    await expect(input).toHaveValue('uncommitted');
    await input.press('ControlOrMeta+A');
    const cdp = await page.context().newCDPSession(page);
    await cdp.send('Input.imeSetComposition', { text: '取消中文', selectionStart: 4, selectionEnd: 4 });
    await expect(input).toHaveValue('取消中文');
    await input.press('Escape');
    await cdp.send('Input.imeSetComposition', { text: '', selectionStart: 0, selectionEnd: 0 });
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(seed);
    expect((await save(page, info, 'cancelled')).cells.A1).toEqual(textCell(seed));
    await expect(page.getByTestId('xlsx-undo')).toBeDisabled();
    await reopen(page, info, [{ address: 'A1', input: seed, cell: textCell(seed) }]);
    await cdp.detach();
  });
}

test('xlsx input: the first keystroke starts a replacing CJK composition', async ({ page }, info) => {
  await open(page);
  await page.getByTestId('xlsx-scroll').press('k');
  const input = page.getByTestId('xlsx-cell-editor');
  await expect(input).toBeFocused();
  await expect(input).toHaveValue('k');
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Input.imeSetComposition', {
    text: 'かな', selectionStart: 2, selectionEnd: 2, replacementStart: 0, replacementEnd: 1,
  });
  await expect(input).toHaveValue('かな');
  await cdp.send('Input.insertText', { text: '仮名' });
  await expect(input).toHaveValue('仮名');
  await input.press('Enter');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A2');
  const entry: Entry = { address: 'A1', input: '仮名', cell: textCell('仮名') };
  await history(page, info, entry);
  await reopen(page, info, [entry]);
  await cdp.detach();
});

const cases: { title: string; entries: Entry[] }[] = [
  {
    title: 'emoji, combining marks, and RTL text retain their code points',
    entries: [
      { address: 'A1', input: 'A😀👩🏽‍💻🚀Z', cell: textCell('A😀👩🏽‍💻🚀Z') },
      { address: 'B1', input: 'Cafe\u0301 a\u0308 n\u0303', cell: textCell('Cafe\u0301 a\u0308 n\u0303'), source: 'formula' },
      { address: 'C1', input: 'שלום مرحبا 123', cell: textCell('שלום مرحبا 123') },
    ],
  },
  {
    title: '5,000 characters survive editing and reopening without truncation',
    entries: [{ address: 'A1', input: 'abcde'.repeat(1000), cell: textCell('abcde'.repeat(1000)) }],
  },
  {
    title: 'apostrophes force numeric and leading-equals text while string formulas stay formulas',
    entries: [
      { address: 'A1', input: "'123", cell: textCell('123') },
      { address: 'B1', input: "'=1+2", cell: textCell('=1+2'), source: 'formula' },
      { address: 'C1', input: '="=literal"', cell: { type: 'str', value: '=literal', formula: '"=literal"' } },
    ],
  },
  {
    title: 'scientific notation stores finite numeric values',
    entries: [
      { address: 'A1', input: '1.25e3', cell: numberCell('1250'), display: /1250|1\.25[Ee]\+?0?3/ },
      { address: 'B1', input: '-2.5E-3', cell: numberCell('-0.0025'), display: /-0\.0025|-2\.5[Ee]-0?3/, source: 'formula' },
    ],
  },
  {
    title: 'TRUE and FALSE are boolean cells regardless of case',
    entries: [
      { address: 'A1', input: 'TRUE', cell: { type: 'b', value: '1', formula: null }, display: 'TRUE' },
      { address: 'B1', input: 'FALSE', cell: { type: 'b', value: '0', formula: null }, display: 'FALSE', source: 'formula' },
      { address: 'C1', input: 'tRuE', cell: { type: 'b', value: '1', formula: null }, display: 'TRUE' },
    ],
  },
  {
    title: 'leading zeros coerce unless an apostrophe forces text',
    entries: [
      { address: 'A1', input: '00123', cell: numberCell('123') },
      { address: 'B1', input: "'00123", cell: textCell('00123'), source: 'formula' },
    ],
  },
  {
    title: 'numeric whitespace trims while text whitespace is preserved',
    entries: [
      { address: 'A1', input: '  42  ', cell: numberCell('42') },
      { address: 'B1', input: '  padded text  ', cell: textCell('  padded text  '), source: 'formula' },
      { address: 'C1', input: '   ', cell: textCell('   ') },
    ],
  },
];

for (const { title, entries } of cases) {
  test('xlsx input: ' + title, async ({ page }, info) => {
    await open(page);
    for (const entry of entries) {
      await select(page, entry.address);
      const input = await startEdit(page, entry.source);
      await page.keyboard.insertText(entry.input);
      await expect(input).toHaveValue(entry.input);
      await input.press('Enter');
      await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
      await expect(page.getByTestId('xlsx-name-box')).toHaveValue(entry.address[0] + '2');
      await history(page, info, entry);
    }
    await reopen(page, info, entries);
  });
}

for (const { title, input, value } of [
  { title: 'thousand separators produce a numeric cell', input: '1,234.5', value: '1234.5' },
  { title: 'a percentage stores its fractional numeric value', input: '12.5%', value: '0.125' },
  { title: 'parentheses produce a negative number', input: '(1234.5)', value: '-1234.5' },
]) {
  test('xlsx input: ' + title, async ({ page }, info) => {
    await open(page);
    const editor = await startEdit(page);
    await page.keyboard.insertText(input);
    await expect(editor).toHaveValue(input);
    await editor.press('Enter');
    await expect(page.getByTestId('xlsx-cell-editor')).toHaveCount(0);
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A2');
    expect((await save(page, info, 'numeric-entry')).cells.A1).toEqual(numberCell(value));
  });
}
