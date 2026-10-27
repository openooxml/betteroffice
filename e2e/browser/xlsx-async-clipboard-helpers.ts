import { expect, type Page } from 'playwright/test';
import JSZip from 'jszip';

export async function clipboardWorkbook() {
  const zip = new JSZip();
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${[1, 2].map((n) => `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`);
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', `<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${[1, 2].map((n) => `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`).join('')}<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><i/><sz val="11"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${ns}"><dimension ref="A1:H8"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="24"/><sheetData><row r="1"><c r="A1" s="1"><v>11</v></c><c r="B1"><v>12</v></c><c r="D1"><f>A1+$B$1+A$2+$B2</f><v>66</v></c><c r="F1" t="inlineStr"><is><t>merged</t></is></c></row><row r="2"><c r="A2"><v>21</v></c><c r="B2"><v>22</v></c></row><row r="8"><c r="H8" t="inlineStr"><is><t>extent</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="F1:G1"/></mergeCells></worksheet>`);
  zip.file('xl/worksheets/sheet2.xml', `<worksheet xmlns="${ns}"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>other</t></is></c></row></sheetData></worksheet>`);
  return { name: 'clipboard.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

export async function workbookXml(page: Page, bytes: Buffer) {
  const zip = await JSZip.loadAsync(bytes);
  const xml: Record<string, string> = {};
  for (const part of Object.values(zip.files)) {
    if (!part.dir && part.name.endsWith('.xml')) xml[part.name] = await part.async('string');
  }
  return page.evaluate((parts) => {
    const parse = (xml: string) => {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      if (doc.querySelector('parsererror')) throw new Error('Invalid saved XML');
      return doc;
    };
    const elements = (node: Document | Element, name: string) => Array.from(node.getElementsByTagNameNS('*', name));
    const text = (node: Document | Element, name: string) => elements(node, name).map((item) => item.textContent ?? '').join('');
    const shared = parts['xl/sharedStrings.xml'] ? elements(parse(parts['xl/sharedStrings.xml']), 'si').map((si) => text(si, 't')) : [];
    const styles = parse(parts['xl/styles.xml']);
    const xfs = elements(elements(styles, 'cellXfs')[0], 'xf');
    const fonts = elements(styles, 'font');
    const fills = elements(styles, 'fill');
    const sheets: Record<string, { cells: Record<string, { value: string; formula: string; bold: boolean; italic: boolean; fill: string }>; merges: string[]; dimension: string | null }> = {};
    for (const [name, xml] of Object.entries(parts)) {
      if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue;
      const doc = parse(xml);
      const cells: typeof sheets[string]['cells'] = {};
      for (const cell of elements(doc, 'c')) {
        const xf = xfs[Number(cell.getAttribute('s') ?? 0)];
        const font = fonts[Number(xf?.getAttribute('fontId') ?? 0)];
        const fill = fills[Number(xf?.getAttribute('fillId') ?? 0)];
        const flag = (name: string) => elements(font, name).some((item) => !['0', 'false'].includes(item.getAttribute('val') ?? '1'));
        const raw = text(cell, 'v');
        cells[cell.getAttribute('r')!] = {
          value: cell.getAttribute('t') === 's' ? shared[Number(raw)] : cell.getAttribute('t') === 'inlineStr' ? text(cell, 't') : raw,
          formula: text(cell, 'f'), bold: flag('b'), italic: flag('i'),
          fill: elements(fill, 'fgColor')[0]?.getAttribute('rgb') ?? '',
        };
      }
      sheets[name] = {
        cells,
        merges: elements(doc, 'mergeCell').map((merge) => merge.getAttribute('ref')!),
        dimension: elements(doc, 'dimension')[0]?.getAttribute('ref') ?? null,
      };
    }
    return sheets;
  }, xml);
}

export async function selectCell(page: Page, row: number, col: number) {
  const scroll = page.getByTestId('xlsx-scroll');
  const nameBox = page.getByTestId('xlsx-name-box');
  // The name box shows the previous selection until the worker applies a move.
  await scroll.press('ControlOrMeta+Home');
  await expect(nameBox).toHaveValue('A1');
  for (let index = 0; index < row; index++) await scroll.press('ArrowDown');
  await expect(nameBox).toHaveValue('A' + (row + 1));
  let column = '';
  for (let index = col + 1; index > 0; index = Math.floor((index - 1) / 26)) column = String.fromCharCode(65 + (index - 1) % 26) + column;
  for (let index = 0; index < col && await nameBox.inputValue() !== column + (row + 1); index++) {
    const from = await nameBox.inputValue();
    await scroll.press('ArrowRight');
    await expect(nameBox).not.toHaveValue(from);
  }
  await expect(nameBox).toHaveValue(column + (row + 1));
}
