import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { implicitFontWorkbook } from './format-fixtures';
import { openXlsx } from '../src/xlsx';
import { registerXlsxTools } from '../src/xlsx-mcp';

export async function chartWorkbook() {
  const zip = await JSZip.loadAsync(await implicitFontWorkbook());
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'));
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', sheet.replace('</worksheet>', '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="d1"/></worksheet>'));
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="d1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
  zip.file('xl/drawings/drawing1.xml', '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><c:chart r:id="c1"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>');
  zip.file('xl/drawings/_rels/drawing1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="c1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>');
  zip.file('xl/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>Sales </a:t></a:r><a:r><a:rPr i="1"/><a:t>report</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strRef><c:f>Sales!$A$2:$A$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Sales!$B$2:$B$4</c:f></c:numRef></c:val></c:ser></c:barChart><c:catAx><c:axId val="7"/></c:catAx></c:plotArea></c:chart></c:chartSpace>');
  const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
  const anchor = drawing.slice(drawing.indexOf('<xdr:twoCellAnchor>'), drawing.indexOf('</xdr:twoCellAnchor>') + '</xdr:twoCellAnchor>'.length).replace('r:id="c1"', 'r:id="c2"');
  zip.file('xl/drawings/drawing1.xml', drawing.replace('</xdr:wsDr>', `${anchor}</xdr:wsDr>`));
  const rels = await zip.file('xl/drawings/_rels/drawing1.xml.rels')!.async('string');
  zip.file('xl/drawings/_rels/drawing1.xml.rels', rels.replace('</Relationships>', '<Relationship Id="c2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>'));
  zip.file('xl/charts/chart2.xml', await zip.file('xl/charts/chart1.xml')!.async('string'));
  const allTypes = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', allTypes.replace('</Types>', '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>'));
  zip.file('xl/charts/untouched.xml', '<untouched a="b" />');
  return zip.generateAsync({ type: 'uint8array' });
}

test('chart MCP reads titles series anchors and edits atomically through save reopen', async () => {
  const source = await chartWorkbook();
  const book = await openXlsx(source);
  const tools = new Map<string, { run: (args: any) => any; readOnly: boolean }>();
  const register: Parameters<typeof registerXlsxTools>[0] = (name, _description, _shape, readOnly, run) => { tools.set(name, { run, readOnly }); };
  try {
    registerXlsxTools(register, { get: () => book });
    expect(tools.get('xlsx_list_charts')?.readOnly).toBe(true);
    expect(tools.get('xlsx_read_chart_text')?.readOnly).toBe(true);
    expect(tools.get('xlsx_edit_charts')?.readOnly).toBe(false);
    const call = (name: string, args: object = {}) => tools.get(name)!.run({ document: 'book', ...args }).structuredContent;
    const sheet = book.outline().items[0].sheetId;
    const listed = call('xlsx_list_charts', { sheet });
    expect(listed.items[0]).toMatchObject({ index: 0, part: 'xl/charts/chart1.xml', type: 'barChart', anchorRange: 'D2:J13', title: 'Sales report', richTitle: true, series: [{ index: 0, name: 'Revenue', nameRef: null, categoryRef: 'Sales!$A$2:$A$4', valueRef: 'Sales!$B$2:$B$4' }] });
    const id = listed.items[0].id;
    expect(() => call('xlsx_edit_charts', { version: listed.version, edits: [{ action: 'set_chart_title', id, title: 'Rollback' }, { action: 'set_series', sheet, chart: 0, series: 8, values: 'Sales!C2:C4' }] })).toThrow('series');
    expect(call('xlsx_list_charts', { sheet })).toEqual(listed);
    const edited = call('xlsx_edit_charts', { version: listed.version, edits: [{ action: 'set_chart_title', id, title: 'Unified title' }, { action: 'set_series', sheet, chart: 0, series: 0, name: { ref: '=Sales!$C$1' }, categories: 'Sales!$A$2:$A$4', values: 'Sales!$C$2:$C$4' }] });
    expect(edited.applied).toBe(true);
    expect(() => call('xlsx_edit_charts', { version: listed.version, edits: [{ action: 'set_chart_title', id, title: null }] })).toThrow('Workbook changed');
    const saved = await book.export();
    const reopened = await openXlsx(saved);
    try {
      const chart = reopened.listCharts({ sheet }).items[0];
      expect(chart).toMatchObject({ id, title: 'Unified title', richTitle: false, series: [{ index: 0, name: 'Cost', nameRef: 'Sales!$C$1', valueRef: 'Sales!$C$2:$C$4' }] });
      const before = await JSZip.loadAsync(source);
      const after = await JSZip.loadAsync(saved);
      for (const path of ['xl/charts/untouched.xml', 'xl/charts/chart2.xml', 'xl/drawings/drawing1.xml']) expect(await after.file(path)!.async('string')).toBe(await before.file(path)!.async('string'));
      const xml = await after.file('xl/charts/chart1.xml')!.async('string');
      expect(xml).toContain('<c:v>50</c:v>');
      expect(xml).toContain('<a:rPr b="1"/>');
      expect(xml.match(/<a:r>/g)).toHaveLength(1);
      call('xlsx_edit_charts', { version: edited.version, edits: [{ action: 'set_series', id, series: 0, name: { text: 'Literal' } }, { action: 'set_chart_title', id, title: null }] });
      expect(call('xlsx_list_charts', { sheet }).items[0]).toMatchObject({ title: null, series: [{ name: 'Literal', nameRef: null }] });
    } finally { reopened.close(); }
    tools.clear();
    registerXlsxTools(register, { get: () => book }, true);
    expect(tools.has('xlsx_edit_charts')).toBe(false);
    expect(tools.get('xlsx_list_charts')?.readOnly).toBe(true);
  } finally { book.close(); }
});

test('chart tools validate MCP schemas and export a reopenable workbook', async () => {
  const { mkdtemp, writeFile, readFile, rm } = await import('node:fs/promises');
  const { tmpdir } = await import('node:os');
  const { join, resolve } = await import('node:path');
  const { Client } = await import('@modelcontextprotocol/sdk/client/index.js');
  const { StdioClientTransport } = await import('@modelcontextprotocol/sdk/client/stdio.js');
  const root = await mkdtemp(join(tmpdir(), 'xlsx-charts-'));
  const client = new Client({ name: 'xlsx-charts-test', version: '1' });
  try {
    await writeFile(join(root, 'book.xlsx'), await chartWorkbook());
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [resolve(import.meta.dir, '../src/cli.ts'), '--root', root], stderr: 'pipe' }));
    const tools = (await client.listTools()).tools;
    expect(tools.find(tool => tool.name === 'xlsx_list_charts')?.annotations?.readOnlyHint).toBe(true);
    expect(tools.find(tool => tool.name === 'xlsx_edit_charts')?.inputSchema.additionalProperties).toBe(false);
    const call = async (name: string, args: object = {}) => {
      const result = await client.callTool({ name, arguments: args as Record<string, unknown> });
      expect(result.isError, JSON.stringify(result.structuredContent)).not.toBe(true);
      return result.structuredContent as Record<string, any>;
    };
    const { document } = await call('office_open', { path: 'book.xlsx' });
    const outline = await call('xlsx_outline', { document });
    const listed = await call('xlsx_list_charts', { document, sheet: outline.items[0].sheetId });
    const id = listed.items[0].id;
    expect(await call('xlsx_read_chart_text', { document, id, field: 'title', length: 5 })).toMatchObject({ text: 'Sales', nextStart: 5, present: true });
    const invalid = await client.callTool({ name: 'xlsx_edit_charts', arguments: { document, version: listed.version, edits: [{ action: 'set_series', id, series: 0, name: { text: 'Name', ref: 'Sales!B1' } }] } });
    expect(invalid.isError).toBe(true);
    const changed = await call('xlsx_edit_charts', { document, version: listed.version, edits: [{ action: 'set_chart_title', id, title: null }, { action: 'set_series', id, series: 0, name: { text: 'New name' } }] });
    expect(changed.applied).toBe(true);
    await call('office_export', { document, path: 'result.xlsx' });
    const reopened = await openXlsx(await readFile(join(root, 'result.xlsx')));
    try { expect(reopened.listCharts({ sheet: outline.items[0].sheetId }).items[0]).toMatchObject({ title: null, richTitle: false, series: [{ name: 'New name', nameRef: null }] }); }
    finally { reopened.close(); }
  } finally { await client.close(); await rm(root, { recursive: true, force: true }); }
}, 30000);

test('chart metadata previews are bounded and complete title and series text remain readable', async () => {
  const book = await openXlsx(await chartWorkbook());
  try {
    const sheet = book.outline().items[0].sheetId;
    const title = 'Title 😀 '.repeat(2000);
    const name = 'Long series '.repeat(600);
    const edited = book.editCharts({ version: book.handle.version(), edits: [{ action: 'set_chart_title', sheet, chart: 0, title }, { action: 'set_series', sheet, chart: 0, series: 0, name: { text: name } }] });
    expect(JSON.stringify(edited).length).toBeLessThan(16000);
    const page = book.listCharts({ sheet, limit: 1 });
    expect(JSON.stringify(page).length).toBeLessThan(16000);
    expect(page.nextOffset).toBe(1);
    expect(page.items[0].truncated).toBe(true);
    expect(page.items[0].title!.length).toBeLessThanOrEqual(400);
    expect(page.items[0].series[0].name!.length).toBeLessThanOrEqual(120);
    for (const [field, expected] of [['title', title], ['name', name]] as const) {
      let text = '';
      let start = 0;
      while (true) {
        const window = book.readChartText({ id: page.items[0].id, field, start });
        expect(JSON.stringify(window).length).toBeLessThan(16000);
        expect(window.present).toBe(true);
        text += window.text;
        if (window.nextStart === null) break;
        start = window.nextStart;
      }
      expect(text).toBe(expected);
    }
    expect(book.readChartText({ id: page.items[0].id, field: 'nameRef' })).toMatchObject({ present: false, text: '', nextStart: null });
    expect(() => book.readChartText({ id: page.items[0].id, field: 'name', series: 99 })).toThrow('series');
  } finally { book.close(); }
});

test('chart parts on chartsheets are editable while their sheet XML and cell write policy stay intact', async () => {
  const zip = await JSZip.loadAsync(await chartWorkbook());
  const workbook = await zip.file('xl/workbook.xml')!.async('string');
  zip.file('xl/workbook.xml', workbook.replace('</sheets>', '<sheet name="Chart" sheetId="2" r:id="rChart"/></sheets>'));
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  zip.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>'));
  const worksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file('xl/worksheets/sheet1.xml', worksheet.replace(/<drawing\b[^>]*\/>/, ''));
  zip.file('xl/chartsheets/_rels/sheet1.xml.rels', await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string'));
  zip.remove('xl/worksheets/_rels/sheet1.xml.rels');
  const chartSheet = '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><drawing r:id="d1"/></chartsheet>';
  zip.file('xl/chartsheets/sheet1.xml', chartSheet);
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/></Types>'));
  const book = await openXlsx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const sheet = book.outline().items.find(sheet => sheet.name === 'Chart')!;
    expect(sheet).toMatchObject({ kind: 'chartsheet', editable: false });
    const listed = book.listCharts({ sheet: sheet.sheetId });
    expect(listed.total).toBe(2);
    book.editCharts({ version: listed.version, edits: [{ action: 'set_chart_title', id: listed.items[0].id, title: 'Chart view' }, { action: 'set_series', sheet: sheet.sheetId, chart: 0, series: 0, name: { text: 'Chart series' }, values: 'Sales!$C$2:$C$4' }] });
    expect(() => book.writeRanges({ version: book.handle.version(), edits: [{ sheet: sheet.sheetId, range: 'A1', values: [[1]] }] })).toThrow();
    const saved = await book.export();
    const parts = await JSZip.loadAsync(saved);
    expect(await parts.file('xl/chartsheets/sheet1.xml')!.async('string')).toBe(chartSheet);
    const reopened = await openXlsx(saved);
    try { expect(reopened.listCharts({ sheet: 'Chart' }).items[0]).toMatchObject({ title: 'Chart view', series: [{ name: 'Chart series', valueRef: 'Sales!$C$2:$C$4' }] }); }
    finally { reopened.close(); }
  } finally { book.close(); }
});

test('unreadable charts stay listed with an error while other charts remain editable', async () => {
  const zip = await JSZip.loadAsync(await chartWorkbook());
  zip.file('xl/charts/chart2.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:roundedCorners val="0"/></c:chartSpace>');
  const book = await openXlsx(await zip.generateAsync({ type: 'uint8array' }));
  try {
    const sheet = book.outline().items[0].sheetId;
    const listed = book.listCharts({ sheet });
    expect(listed.items.map(item => [item.editable, item.error === null])).toEqual([[true, true], [false, false]]);
    expect(listed.items[1].error).toContain('no chart element');
    expect(listed.items[0].autoTitle).toBe(false);
    expect(() => book.editCharts({ version: listed.version, edits: [{ action: 'set_chart_title', id: listed.items[1].id, title: 'No' }] })).toThrow('no chart element');
    expect(() => book.readChartText({ id: listed.items[1].id })).toThrow('no chart element');
    expect(book.listCharts({ sheet })).toEqual(listed);
    const edited = book.editCharts({ version: listed.version, edits: [{ action: 'set_chart_title', id: listed.items[0].id, title: 'Readable' }] });
    expect(edited.applied).toBe(true);
    expect(book.listCharts({ sheet }).items[0].title).toBe('Readable');
  } finally { book.close(); }
});
