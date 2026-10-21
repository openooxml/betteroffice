import { test, expect, type Locator, type Page, type TestInfo } from 'playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rangeCells = ['B2', 'C2', 'B3', 'C3'];
const targets = [['A1'], rangeCells];

type Edge = { style: string | null; color: string | null };
type CellStyle = {
  bold: boolean;
  italic: boolean;
  strike: boolean;
  family: string;
  size: number;
  textColor: string | null;
  fillColor: string | null;
  numberFormat: string;
  horizontal: string;
  vertical: string;
  wrap: boolean;
  shrink: boolean;
  borders: Record<string, Edge>;
};
type Saved = {
  file: string;
  cells: Record<string, CellStyle>;
  values: Record<string, string>;
  merges: string[];
};

const rail = (page: Page) => page.getByTestId('xlsx-formatting-toolbar');
const button = (page: Page, name: string) => rail(page).getByRole('button', { name, exact: true });

async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    </Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  </Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <bookViews><workbookView activeTab="0"/></bookViews>
    <sheets><sheet name="Formatting" sheetId="1" r:id="rId1"/></sheets>
  </workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${ns}">
    <fonts count="1"><font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font></fonts>
    <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
    <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
    <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  </styleSheet>`);
  const rows = Array.from({ length: 8 }, (_, row) => `<row r="${row + 1}">${
    Array.from({ length: 8 }, (_, col) => {
      const address = `${String.fromCharCode(65 + col)}${row + 1}`;
      if (row === 0 && col === 0) return `<c r="${address}" t="inlineStr"><is><t>Format me</t></is></c>`;
      if (address === 'H8') return `<c r="${address}"><v>1</v></c>`;
      return row < 4 ? `<c r="${address}"><v>12.3456</v></c>` : `<c r="${address}"/>`;
    }).join('')
  }</row>`).join('');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${ns}">
    <dimension ref="A1:H8"/>
    <sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>
    <sheetFormatPr defaultRowHeight="24" defaultColWidth="14"/>
    <cols><col min="1" max="8" width="14" customWidth="1"/></cols>
    <sheetData>${rows}</sheetData>
  </worksheet>`);
  return { name: 'formatting.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

async function open(page: Page, file?: string) {
  await page.goto('/?format=xlsx');
  await page.locator('input[type=file]').last().setInputFiles(file ?? await fixture());
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  await expect(button(page, 'Bold')).toBeEnabled();
  await expect(page.getByTestId('xlsx-formula-input')).toBeEditable();
}

async function select(page: Page, address: string) {
  const match = /^([A-H])([1-8])$/.exec(address)!;
  const grid = page.getByTestId('xlsx-scroll');
  await grid.press('ControlOrMeta+Home');
  for (let col = 0; col < match[1].charCodeAt(0) - 65; col++) await grid.press('ArrowRight');
  for (let row = 1; row < Number(match[2]); row++) await grid.press('ArrowDown');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue(address);
}

async function selectRange(page: Page, start = 'B2', end = 'C3', drag = false) {
  if (drag) {
    await select(page, end);
    const finish = (await page.getByTestId('xlsx-selection').boundingBox())!;
    await select(page, start);
    const begin = (await page.getByTestId('xlsx-selection').boundingBox())!;
    await page.mouse.move(begin.x + begin.width / 2, begin.y + begin.height / 2);
    await page.mouse.down();
    await page.mouse.move(finish.x + finish.width / 2, finish.y + finish.height / 2, { steps: 4 });
    await page.mouse.up();
  } else {
    await select(page, start);
    const grid = page.getByTestId('xlsx-scroll');
    for (let col = start.charCodeAt(0); col < end.charCodeAt(0); col++) await grid.press('Shift+ArrowRight');
    for (let row = Number(start.slice(1)); row < Number(end.slice(1)); row++) await grid.press('Shift+ArrowDown');
  }
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue(start);
  await expect(page.getByRole('gridcell', { selected: true })).toHaveCount(
    (end.charCodeAt(0) - start.charCodeAt(0) + 1) * (Number(end.slice(1)) - Number(start.slice(1)) + 1)
  );
}

async function target(page: Page, cells: string[], drag = false) {
  if (cells.length === 1) await select(page, cells[0]);
  else await selectRange(page, cells[0], cells.at(-1)!, drag);
}

async function choose(page: Page, control: string, option: string) {
  await button(page, control).click();
  await page.getByRole('menu', { name: control, exact: true })
    .getByRole(control === 'Merge cells' ? 'menuitem' : 'menuitemradio', { name: option, exact: true }).click();
}

async function checked(page: Page, control: string, option: string, active = true) {
  await button(page, control).click();
  const menu = page.getByRole('menu', { name: control, exact: true });
  await expect(menu.getByRole('menuitemradio', { name: option, exact: true }))
    .toHaveAttribute('aria-checked', String(active));
  await menu.press('Escape');
}

async function color(input: Locator, value: string) {
  // Native color dialogs live outside the page; emit the input's platform change events.
  await input.evaluate((element, value) => {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
    setter.call(element, value);
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
  }, value);
}

async function save(page: Page, info: TestInfo, label: string): Promise<Saved> {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const download = await downloading;
  const file = info.outputPath(label + '.xlsx');
  await download.saveAs(file);
  const zip = await JSZip.loadAsync(await readFile(file));
  const styles = await zip.file('xl/styles.xml')!.async('string');
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const sharedStrings = await zip.file('xl/sharedStrings.xml')?.async('string') ?? '';
  const result = await page.evaluate(({ styles, sheet, sharedStrings }) => {
    const parse = (xml: string) => {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      if (doc.querySelector('parsererror')) throw new Error('Invalid saved XML');
      return doc;
    };
    const styleDoc = parse(styles);
    const sheetDoc = parse(sheet);
    const children = (node: ParentNode | undefined, name: string): Element[] =>
      Array.from(node?.children ?? []).filter((child) => child.localName === name);
    const child = (node: ParentNode | undefined, name: string) => children(node, name)[0];
    const table = (name: string, item: string) => children(child(styleDoc.documentElement, name), item);
    const fonts = table('fonts', 'font');
    const fills = table('fills', 'fill');
    const borders = table('borders', 'border');
    const xfs = table('cellXfs', 'xf');
    const custom = new Map(table('numFmts', 'numFmt').map((node) =>
      [node.getAttribute('numFmtId'), node.getAttribute('formatCode')!]));
    const builtins: Record<string, string> = {
      0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
      9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'mm-dd-yy', 20: 'h:mm',
      21: 'h:mm:ss', 49: '@',
    };
    const rgb = (node?: Element) => node?.getAttribute('rgb')?.slice(-6).toUpperCase() ?? null;
    const flag = (node: Element | undefined, name: string) => {
      const value = child(node, name);
      return !!value && !['0', 'false'].includes(value.getAttribute('val') ?? '1');
    };
    const strings = sharedStrings ? Array.from(parse(sharedStrings).getElementsByTagNameNS('*', 'si'))
      .map((si) => Array.from(si.getElementsByTagNameNS('*', 't')).map((t) => t.textContent ?? '').join('')) : [];
    const cells: Record<string, CellStyle> = {};
    const values: Record<string, string> = {};
    const nodes = Array.from(sheetDoc.getElementsByTagNameNS('*', 'c'));
    for (let row = 1; row <= 8; row++) {
      for (let col = 0; col < 8; col++) {
        const address = `${String.fromCharCode(65 + col)}${row}`;
        const cell = nodes.find((node) => node.getAttribute('r') === address);
        const index = Number(cell?.getAttribute('s') ?? 0);
        const xf = xfs[index];
        if (!xf) throw new Error(`Missing cellXf ${index} for ${address}`);
        const font = fonts[Number(xf.getAttribute('fontId') ?? 0)];
        const fill = fills[Number(xf.getAttribute('fillId') ?? 0)];
        const border = borders[Number(xf.getAttribute('borderId') ?? 0)];
        if (!font || !fill || !border) throw new Error(`Invalid style references for ${address}`);
        const alignment = child(xf, 'alignment');
        const pattern = child(fill, 'patternFill');
        const numFmtId = xf.getAttribute('numFmtId') ?? '0';
        cells[address] = {
          bold: flag(font, 'b'), italic: flag(font, 'i'), strike: flag(font, 'strike'),
          family: child(font, 'name')?.getAttribute('val') ?? '',
          size: Number(child(font, 'sz')?.getAttribute('val') ?? 0),
          textColor: rgb(child(font, 'color')),
          fillColor: pattern?.getAttribute('patternType') === 'solid' ? rgb(child(pattern, 'fgColor')) : null,
          numberFormat: custom.get(numFmtId) ?? builtins[numFmtId] ?? `builtin:${numFmtId}`,
          horizontal: alignment?.getAttribute('horizontal') ?? 'general',
          vertical: alignment?.getAttribute('vertical') ?? 'bottom',
          wrap: ['1', 'true'].includes(alignment?.getAttribute('wrapText') ?? ''),
          shrink: ['1', 'true'].includes(alignment?.getAttribute('shrinkToFit') ?? ''),
          borders: Object.fromEntries(['left', 'right', 'top', 'bottom'].map((edge) => {
            const side = child(border, edge);
            const style = side?.getAttribute('style') ?? null;
            return [edge, { style, color: style ? rgb(child(side, 'color')) ?? '000000' : null }];
          })),
        };
        const value = child(cell, 'v')?.textContent ?? '';
        values[address] = cell?.getAttribute('t') === 's' ? strings[Number(value)] :
          cell?.getAttribute('t') === 'inlineStr' ? Array.from(cell.getElementsByTagNameNS('*', 't'))
            .map((t) => t.textContent ?? '').join('') : value;
      }
    }
    return { cells, values, merges: Array.from(sheetDoc.getElementsByTagNameNS('*', 'mergeCell'))
      .map((node) => node.getAttribute('ref')!).sort() };
  }, { styles, sheet, sharedStrings });
  return { file, ...result };
}

function sameWorkbook(actual: Saved, expected: Saved) {
  expect(actual.cells).toEqual(expected.cells);
  expect(actual.values).toEqual(expected.values);
  expect(actual.merges).toEqual(expected.merges);
}

function styles(saved: Saved, cells: string[], expected: Partial<CellStyle>) {
  for (const address of cells) expect(saved.cells[address], address).toMatchObject(expected);
}

async function checkpoint(page: Page, info: TestInfo, label: string, before: Saved | undefined,
  visible: () => Promise<void>, undone?: () => Promise<void>) {
  await visible();
  const applied = await save(page, info, label + '-applied');
  if (!before) return applied;
  await page.getByTestId('xlsx-undo').click();
  if (undone) await undone();
  sameWorkbook(await save(page, info, label + '-undo'), before);
  await page.getByTestId('xlsx-redo').click();
  await visible();
  sameWorkbook(await save(page, info, label + '-redo'), applied);
  return applied;
}

async function reopen(page: Page, info: TestInfo, saved: Saved, visible?: () => Promise<void>) {
  await open(page, saved.file);
  if (visible) await visible();
  sameWorkbook(await save(page, info, 'reopened'), saved);
}

test.beforeEach(async ({ page }) => {
  await page.setViewportSize({ width: 1920, height: 1080 });
  await open(page);
  await expect(page.getByTestId('xlsx-formula-input')).toHaveValue('Format me');
});

for (const [label, property, shortcut] of [
  ['Bold', 'bold', 'ControlOrMeta+b'],
  ['Italic', 'italic', 'ControlOrMeta+i'],
  ['Strikethrough', 'strike', null],
] as const) {
  test(`xlsx async format: ${label.toLowerCase()} toggles on and off on a cell and range`, async ({ page }, info) => {
    let saved = await save(page, info, 'initial');
    for (const cells of targets) {
      await target(page, cells);
      const state = (active: boolean) => async () => {
        await expect(button(page, label)).toHaveAttribute('aria-pressed', String(active));
      };
      await state(false)();
      if (shortcut) await page.getByTestId('xlsx-scroll').press(shortcut);
      else await button(page, label).click();
      saved = await checkpoint(page, info, cells[0] + '-on', cells === targets[0] ? saved : undefined, state(true), state(false));
      styles(saved, cells, { [property]: true });
      if (cells.length === 1) await reopen(page, info, saved, state(true));
      await select(page, 'H4');
      await state(false)();
      await target(page, cells);
      await state(true)();
      await button(page, label).click();
      saved = await checkpoint(page, info, cells[0] + '-off', undefined, state(false));
      styles(saved, cells, { [property]: false });
    }
    await reopen(page, info, saved, async () => {
      await expect(button(page, label)).toHaveAttribute('aria-pressed', 'false');
    });
  });
}

test('xlsx async format: font family follows the selected cell and dragged range', async ({ page }, info) => {
  let saved = await save(page, info, 'initial');
  for (const cells of targets) {
    await target(page, cells, true);
    await choose(page, 'Font family', 'Georgia');
    saved = await checkpoint(page, info, cells[0], cells === targets[0] ? saved : undefined, async () => {
      await expect(button(page, 'Font family')).toHaveText('Georgia');
      await checked(page, 'Font family', 'Georgia');
    }, async () => { await expect(button(page, 'Font family')).toHaveText('Calibri'); });
    styles(saved, cells, { family: 'Georgia' });
    await select(page, 'H4');
    await expect(button(page, 'Font family')).toHaveText('Calibri');
    await target(page, cells);
    await expect(button(page, 'Font family')).toHaveText('Georgia');
  }
  await reopen(page, info, saved, async () => {
    await expect(button(page, 'Font family')).toHaveText('Georgia');
  });
});

test('xlsx async format: font size entry and size steps preserve cell and range styles', async ({ page }, info) => {
  let saved = await save(page, info, 'initial');
  const size = rail(page).getByRole('combobox', { name: 'Font size', exact: true });
  for (const cells of targets) {
    await target(page, cells);
    for (const [action, expected, previous] of [['entry', 14, 11], ['increase', 16, 14], ['decrease', 14, 16]] as const) {
      if (action === 'entry') {
        await size.fill('14');
        await size.press('Enter');
      } else await button(page, action === 'increase' ? 'Increase font size' : 'Decrease font size').click();
      saved = await checkpoint(page, info, cells[0] + action, cells === targets[0] && action === 'entry' ? saved : undefined,
        async () => { await expect(size).toHaveValue(String(expected)); },
        async () => { await expect(size).toHaveValue(String(previous)); });
      styles(saved, cells, { size: expected });
    }
    await select(page, 'H4');
    await expect(size).toHaveValue('11');
    await target(page, cells);
    await expect(size).toHaveValue('14');
  }
  await reopen(page, info, saved, async () => { await expect(size).toHaveValue('14'); });
});

for (const [label, property, value, original] of [
  ['Text color', 'textColor', '#c00000', '#000000'],
  ['Fill color', 'fillColor', '#ffff00', '#ffffff'],
] as const) {
  test(`xlsx async format: ${label.toLowerCase()} follows a cell and dragged range`, async ({ page }, info) => {
    let saved = await save(page, info, 'initial');
    const input = rail(page).locator(`input[type=color][aria-label="${label}"]`);
    for (const cells of targets) {
      await target(page, cells, true);
      await color(input, value);
      saved = await checkpoint(page, info, cells[0], cells === targets[0] ? saved : undefined,
        async () => { await expect(input).toHaveValue(value); },
        async () => { await expect(input).toHaveValue(original); });
      styles(saved, cells, { [property]: value.slice(1).toUpperCase() });
      await select(page, 'H4');
      await expect(input).toHaveValue(original);
      await target(page, cells);
      await expect(input).toHaveValue(value);
    }
    await reopen(page, info, saved, async () => { await expect(input).toHaveValue(value); });
  });
}

for (const cells of [['A2'], rangeCells]) {
  for (const [facet, choices] of [
    ['number formats', [
      ['Number', '#,##0.00', null],
      ['Scientific', '0.00E+00', null],
      ['Date', 'mm-dd-yy', null],
      ['Time', 'h:mm', null],
      ['Automatic', 'General', null],
      ['Plain text', '@', null],
    ]],
    ['currency and percent shortcuts', [
      ['Percent', '0.00%', 'Format as percent'],
      ['Currency', null, 'Format as currency'],
    ]],
  ] as const) {
    test(`xlsx async format: ${facet} preserve numeric values on ${cells.join(':')}`, async ({ page }, info) => {
      let saved = await save(page, info, 'initial');
      await target(page, cells);
      for (const [index, [label, pattern, shortcut]] of choices.entries()) {
        if (shortcut) await button(page, shortcut).click();
        else await choose(page, 'More number formats', label);
        saved = await checkpoint(page, info, label, index === 0 ? saved : undefined,
          () => checked(page, 'More number formats', label),
          () => checked(page, 'More number formats', 'Automatic'));
        for (const address of cells) {
          if (pattern) expect(saved.cells[address].numberFormat).toBe(pattern);
          else expect(saved.cells[address].numberFormat).toMatch(/\$|¤|builtin:[5-8]$/);
          expect(saved.values[address]).toBe('12.3456');
        }
      }
      const last = choices.at(-1)![0];
      await select(page, 'H4');
      await checked(page, 'More number formats', 'Automatic');
      await target(page, cells);
      await checked(page, 'More number formats', last);
      await reopen(page, info, saved, async () => {
        await target(page, cells);
        await checked(page, 'More number formats', last);
      });
    });
  }
}

test('xlsx async format: decimal places increase and decrease on a cell and range', async ({ page }, info) => {
  let saved = await save(page, info, 'initial');
  for (const cells of [['A2'], rangeCells]) {
    await target(page, cells);
    await choose(page, 'More number formats', 'Number');
    await checked(page, 'More number formats', 'Number');
    saved = await save(page, info, cells[0] + '-number');
    for (const [label, expected, previous] of [
      ['Increase decimal places', '12.346', '12.35'],
      ['Decrease decimal places', '12.35', '12.346'],
    ] as const) {
      await button(page, label).click();
      const display = async (value: string) => {
        await expect(page.getByRole('gridcell', { selected: true })).toHaveCount(cells.length);
        for (const cell of await page.getByRole('gridcell', { selected: true }).all()) {
          await expect(cell).toContainText(value);
        }
      };
      saved = await checkpoint(page, info, cells[0] + label, cells[0] === 'A2' && label === 'Increase decimal places' ? saved : undefined, () => display(expected), () => display(previous));
      styles(saved, cells, { numberFormat: expected === '12.346' ? '#,##0.000' : '#,##0.00' });
      for (const address of cells) expect(saved.values[address]).toBe('12.3456');
    }
  }
  await reopen(page, info, saved, async () => {
    await select(page, 'A2');
    await expect(page.getByRole('gridcell', { selected: true })).toContainText('12.35');
  });
});

async function pickBorderColor(page: Page, value: string) {
  await button(page, 'Borders').click();
  await page.getByRole('menu', { name: 'Borders', exact: true })
    .getByRole('menuitem', { name: 'Border color', exact: true }).click();
  await color(rail(page).getByRole('group', { name: 'Borders', exact: true })
    .locator('input[type=color][aria-hidden=true]'), value);
  await page.keyboard.press('Escape');
}

async function borderColorState(page: Page, value: string) {
  await button(page, 'Borders').click();
  await expect(page.getByRole('menu', { name: 'Borders', exact: true })
    .getByRole('menuitem', { name: 'Border color', exact: true }).locator('span[aria-hidden=true]'))
    .toHaveCSS('background-color', value);
  await page.getByRole('menu', { name: 'Borders', exact: true }).press('Escape');
}

const edges = ['left', 'right', 'top', 'bottom'];
const allEdges = (style: string, color: string) => Object.fromEntries(
  edges.map((edge) => [edge, { style, color }])
);

for (const cells of targets) {
  test(`xlsx async format: border styles and color apply to ${cells.join(':')}`, async ({ page }, info) => {
    await target(page, cells, true);
    await choose(page, 'Borders', 'Solid border');
    await pickBorderColor(page, '#000000');
    await choose(page, 'Borders', 'All borders');
    await checked(page, 'Borders', 'All borders');
    let saved = await save(page, info, 'all');
    styles(saved, cells, { borders: allEdges('thin', '000000') });
    for (const [label, style] of [
      ['Dashed border', 'dashed'], ['Dotted border', 'dotted'],
      ['Double border', 'double'], ['Solid border', 'thin'],
    ]) {
      await choose(page, 'Borders', label);
      saved = await checkpoint(page, info, label, undefined, () => checked(page, 'Borders', label));
      styles(saved, cells, { borders: allEdges(style, '000000') });
    }
    await pickBorderColor(page, '#0070c0');
    saved = await checkpoint(page, info, 'color', saved,
      () => borderColorState(page, 'rgb(0, 112, 192)'),
      () => borderColorState(page, 'rgb(0, 112, 192)'));
    styles(saved, cells, { borders: allEdges('thin', '0070C0') });
    await select(page, 'H4');
    await checked(page, 'Borders', 'Clear borders');
    await checked(page, 'Borders', 'Solid border');
    await borderColorState(page, 'rgb(0, 112, 192)');
    await reopen(page, info, saved);
    await target(page, cells);
    await borderColorState(page, 'rgb(0, 112, 192)');
    await choose(page, 'Borders', 'Clear borders');
    saved = await checkpoint(page, info, 'clear', undefined, () => checked(page, 'Borders', 'Clear borders'));
    for (const address of cells) for (const edge of Object.values(saved.cells[address].borders)) {
      expect(edge.style).toBeNull();
    }
    await reopen(page, info, saved);
  });

  for (const [facet, presets] of [
    ['geometry', [
      ['outer', 'Outer borders'], ['inner', 'Inner borders'],
      ['horizontal', 'Horizontal borders'], ['vertical', 'Vertical borders'],
    ]],
    ['sides', [
      ['left', 'Left border'], ['top', 'Top border'],
      ['right', 'Right border'], ['bottom', 'Bottom border'],
    ]],
  ] as const) {
    test(`xlsx async format: border ${facet} applies to ${cells.join(':')}`, async ({ page }, info) => {
      const initial = await save(page, info, 'initial');
      let saved = initial;
      await target(page, cells, true);
      await choose(page, 'Borders', 'Solid border');
      for (const [index, [preset, label]] of presets.entries()) {
        await choose(page, 'Borders', 'Clear borders');
        await checked(page, 'Borders', 'Clear borders');
        await pickBorderColor(page, '#0070c0');
        await choose(page, 'Borders', label);
        const empty = cells.length === 1 && ['inner', 'horizontal', 'vertical'].includes(preset);
        const selected = empty ? 'Clear borders' : cells.length === 1 && preset === 'outer' ? 'All borders' : label;
        saved = await checkpoint(page, info, preset, index === 0 ? initial : undefined,
          () => checked(page, 'Borders', selected), () => checked(page, 'Borders', 'Clear borders'));
        if (empty) sameWorkbook(saved, initial);
        for (const address of cells) {
          for (const edge of edges) {
            const boundary = cells.length === 1 || (edge === 'left' ? address.startsWith('B') :
              edge === 'right' ? address.startsWith('C') : edge === 'top' ? address.endsWith('2') : address.endsWith('3'));
            const expected = preset === 'outer' ? boundary : preset === 'inner' ? !boundary :
              preset === 'horizontal' ? ['top', 'bottom'].includes(edge) && !boundary :
              preset === 'vertical' ? ['left', 'right'].includes(edge) && !boundary : preset === edge && boundary;
            expect(saved.cells[address].borders[edge], address + ' ' + edge).toEqual(expected
              ? { style: 'thin', color: '0070C0' } : { style: null, color: null });
          }
        }
      }
      await choose(page, 'Borders', 'Clear borders');
      saved = await checkpoint(page, info, 'clear', undefined, () => checked(page, 'Borders', 'Clear borders'));
      sameWorkbook(saved, initial);
      await reopen(page, info, saved);
    });
  }
}

for (const [action, expected] of [
  ['Merge all', ['D5:E6']],
  ['Merge horizontally', ['D5:E5', 'D6:E6']],
  ['Merge vertically', ['D5:D6', 'E5:E6']],
] as const) {
  test(`xlsx async format: ${action.toLowerCase()} can be undone and unmerged`, async ({ page }, info) => {
    let saved = await save(page, info, 'initial');
    await expect(page.getByTestId('xlsx-merge-all')).toBeDisabled();
    await selectRange(page, 'D5', 'E6', action === 'Merge all');
    const before = (await page.getByTestId('xlsx-selection').boundingBox())!;
    if (action === 'Merge all') await page.getByTestId('xlsx-merge-all').click();
    else await choose(page, 'Merge cells', action);
    saved = await checkpoint(page, info, action, saved, async () => {
      await button(page, 'Merge cells').click();
      await expect(page.getByRole('menu', { name: 'Merge cells', exact: true })
        .getByRole('menuitem', { name: 'Unmerge', exact: true })).not.toHaveAttribute('aria-disabled', 'true');
      await page.getByRole('menu', { name: 'Merge cells', exact: true }).press('Escape');
    });
    expect(saved.merges).toEqual(expected);
    if (action === 'Merge all') {
      await select(page, 'D5');
      const merged = (await page.getByTestId('xlsx-selection').boundingBox())!;
      expect(merged.width).toBeCloseTo(before.width, 0);
      expect(merged.height).toBeCloseTo(before.height, 0);
    }
    await reopen(page, info, saved);
    await selectRange(page, 'D5', 'E6');
    await choose(page, 'Merge cells', 'Unmerge');
    saved = await checkpoint(page, info, action + '-unmerge', undefined, async () => {
      await button(page, 'Merge cells').click();
      await expect(page.getByRole('menu', { name: 'Merge cells', exact: true })
        .getByRole('menuitem', { name: 'Unmerge', exact: true })).toHaveAttribute('aria-disabled', 'true');
      await page.getByRole('menu', { name: 'Merge cells', exact: true }).press('Escape');
    });
    expect(saved.merges).toEqual([]);
    await reopen(page, info, saved);
  });
}

for (const [control, property, choices, original] of [
  ['Horizontal alignment', 'horizontal', [['Center', 'center'], ['Right', 'right'], ['Left', 'left']], 'Left'],
  ['Vertical alignment', 'vertical', [['Top', 'top'], ['Middle', 'center'], ['Bottom', 'bottom']], 'Bottom'],
] as const) {
  for (const cells of targets) {
    test(`xlsx async format: ${control.toLowerCase()} follows ${cells.join(':')}`, async ({ page }, info) => {
      let saved = await save(page, info, 'initial');
      await target(page, cells);
      for (const [index, [label, value]] of choices.entries()) {
        await choose(page, control, label);
        saved = await checkpoint(page, info, label, index === 0 ? saved : undefined,
          () => checked(page, control, label), () => checked(page, control, original));
        styles(saved, cells, { [property]: value });
        if (index === 0) {
          await select(page, 'H4');
          await checked(page, control, original);
          await target(page, cells);
          await checked(page, control, label);
        }
      }
      await select(page, 'H4');
      await checked(page, control, original);
      await target(page, cells);
      await checked(page, control, original);
      await reopen(page, info, saved, () => checked(page, control, original));
    });
  }
}

test('xlsx async format: wrap and overflow persist on a cell and dragged range', async ({ page }, info) => {
  let saved = await save(page, info, 'initial');
  for (const cells of targets) {
    await target(page, cells, true);
    for (const label of ['Wrap', 'Overflow']) {
      await choose(page, 'Text wrapping', label);
      saved = await checkpoint(page, info, cells[0] + label, cells === targets[0] && label === 'Wrap' ? saved : undefined,
        () => checked(page, 'Text wrapping', label), () => checked(page, 'Text wrapping', 'Overflow'));
      styles(saved, cells, { wrap: label === 'Wrap', shrink: false });
    }
    await select(page, 'H4');
    await checked(page, 'Text wrapping', 'Overflow');
    await target(page, cells);
    await checked(page, 'Text wrapping', 'Overflow');
  }
  await reopen(page, info, saved, () => checked(page, 'Text wrapping', 'Overflow'));
});

test('xlsx async format: paint format copies styles to a cell and dragged range without changing values', async ({ page }, info) => {
  await button(page, 'Bold').click();
  await expect(button(page, 'Bold')).toHaveAttribute('aria-pressed', 'true');
  await choose(page, 'Font family', 'Georgia');
  await choose(page, 'Horizontal alignment', 'Center');
  await color(rail(page).locator('input[type=color][aria-label="Fill color"]'), '#ffff00');
  await expect(rail(page).locator('input[type=color][aria-label="Fill color"]')).toHaveValue('#ffff00');
  let saved = await save(page, info, 'source');
  for (const cells of [['D2'], rangeCells]) {
    await select(page, cells.at(-1)!);
    const finish = (await page.getByTestId('xlsx-selection').boundingBox())!;
    await select(page, cells[0]);
    const begin = (await page.getByTestId('xlsx-selection').boundingBox())!;
    await select(page, 'A1');
    await button(page, 'Paint format').click();
    await expect(button(page, 'Paint format')).toHaveAttribute('aria-pressed', 'true');
    await page.mouse.move(begin.x + begin.width / 2, begin.y + begin.height / 2);
    await page.mouse.down();
    await page.mouse.move(finish.x + finish.width / 2, finish.y + finish.height / 2, { steps: 4 });
    await page.mouse.up();
    await expect(page.getByTestId('xlsx-name-box')).toHaveValue(cells[0]);
    saved = await checkpoint(page, info, cells[0], cells[0] === 'D2' ? saved : undefined, async () => {
      await expect(button(page, 'Paint format')).toHaveAttribute('aria-pressed', 'false');
      await expect(button(page, 'Bold')).toHaveAttribute('aria-pressed', 'true');
      await expect(button(page, 'Font family')).toHaveText('Georgia');
      await checked(page, 'Horizontal alignment', 'Center');
    }, async () => { await expect(button(page, 'Bold')).toHaveAttribute('aria-pressed', 'false'); });
    for (const address of cells) {
      expect(saved.cells[address]).toEqual(saved.cells.A1);
      expect(saved.values[address]).toBe('12.3456');
    }
  }
  await select(page, 'A1');
  await button(page, 'Paint format').click();
  await expect(button(page, 'Paint format')).toHaveAttribute('aria-pressed', 'true');
  await button(page, 'Paint format').click();
  await expect(button(page, 'Paint format')).toHaveAttribute('aria-pressed', 'false');
  await select(page, 'H4');
  sameWorkbook(await save(page, info, 'cancelled'), saved);
  await reopen(page, info, saved, async () => {
    await select(page, 'D2');
    await expect(button(page, 'Bold')).toHaveAttribute('aria-pressed', 'true');
    await expect(button(page, 'Font family')).toHaveText('Georgia');
  });
});

for (const [label, property, text] of [
  ['Bold', 'bold', 'Formatted draft'],
  ['Italic', 'italic', 'Formatted italic draft'],
] as const) {
  test(`xlsx async format: ${label.toLowerCase()} during cell edit commits the draft before formatting`, async ({ page }, info) => {
    await page.getByTestId('xlsx-scroll').press('F2');
    const input = page.getByTestId('xlsx-cell-editor');
    await input.fill(text);
    await button(page, label).click();
    await expect(input).toHaveCount(0);
    await expect(button(page, label)).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(text);
    await expect(page.getByRole('gridcell', { selected: true })).toContainText(text);
    const saved = await save(page, info, label + '-draft');
    styles(saved, ['A1'], { [property]: true });
    expect(saved.values.A1).toBe(text);
    await page.getByTestId('xlsx-undo').click();
    await expect(button(page, label)).toHaveAttribute('aria-pressed', 'false');
    const undone = await save(page, info, label + '-draft-undo');
    styles(undone, ['A1'], { [property]: false });
    expect(undone.values.A1).toBe(text);
    await page.getByTestId('xlsx-redo').click();
    await expect(button(page, label)).toHaveAttribute('aria-pressed', 'true');
    sameWorkbook(await save(page, info, label + '-draft-redo'), saved);
    await reopen(page, info, saved, async () => {
      await expect(button(page, label)).toHaveAttribute('aria-pressed', 'true');
      await expect(page.getByTestId('xlsx-formula-input')).toHaveValue(text);
    });
  });
}
