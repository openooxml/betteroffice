import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function fixture() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${relationships}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${relationships}/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>original</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>original</t></is></c></row></sheetData></worksheet>`);
  return { name: 'host-color.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

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

async function editColors(page: Page, host: boolean, cell: 'A1' | 'B1') {
  await page.goto('/?format=xlsx');
  if (host) await page.addStyleTag({ content: 'body { color: rgb(237, 237, 237); } input, textarea { color: inherit; }' });
  await page.locator('input[type=file]').last().setInputFiles(await fixture());
  await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue('A1');
  const scroll = page.getByTestId('xlsx-scroll');
  if (cell === 'B1') await scroll.press('ArrowRight');
  await expect(page.getByTestId('xlsx-name-box')).toHaveValue(cell);
  await scroll.press('F2');
  const input = page.getByTestId('xlsx-cell-editor');
  await expect(input).toHaveValue('original');
  return input.evaluate((element) => {
    const style = getComputedStyle(element);
    return { text: style.color, caret: style.caretColor === 'auto' ? style.color : style.caretColor, background: style.backgroundColor };
  });
}

test('xlsx cell editor text and caret stay dark on a white cell inside a dark-themed host', async ({ page }) => {
  const colors = await editColors(page, true, 'A1');
  expect(colors.background).toBe('rgb(255, 255, 255)');
  expect(ratio(colors.text, colors.background)).toBeGreaterThanOrEqual(3);
  expect(ratio(colors.caret, colors.background)).toBeGreaterThanOrEqual(3);
});

for (const host of [false, true]) {
  test(`xlsx cell editor keeps white text and caret on a dark-filled cell${host ? ' inside a dark-themed host' : ''}`, async ({ page }) => {
    const colors = await editColors(page, host, 'B1');
    expect(colors).toEqual({ text: 'rgb(255, 255, 255)', caret: 'rgb(255, 255, 255)', background: 'rgb(31, 56, 100)' });
  });
}
