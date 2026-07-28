import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(root, 'apps/demo/public/showcase.xlsx');
const zipDate = new Date('2026-01-01T00:00:00Z');

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const sharedStrings: string[] = [];
const sharedStringIds = new Map<string, number>();
let sharedStringUses = 0;

function sharedString(value: string): number {
  sharedStringUses += 1;
  const existing = sharedStringIds.get(value);
  if (existing !== undefined) return existing;
  const id = sharedStrings.length;
  sharedStrings.push(value);
  sharedStringIds.set(value, id);
  return id;
}

const style = {
  title: 1,
  subtitle: 2,
  header: 3,
  currency: 4,
  date: 5,
  percent: 6,
  totalLabel: 7,
  totalCurrency: 8,
  totalPlain: 9,
  callout: 10,
  label: 11,
} as const;

function styleAttribute(value?: number): string {
  return value === undefined ? '' : ` s="${value}"`;
}

function textCell(ref: string, value: string, cellStyle?: number): string {
  return `<c r="${ref}" t="s"${styleAttribute(cellStyle)}><v>${sharedString(value)}</v></c>`;
}

function numberCell(ref: string, value: number, cellStyle?: number): string {
  return `<c r="${ref}"${styleAttribute(cellStyle)}><v>${value}</v></c>`;
}

function formulaCell(ref: string, formula: string, cellStyle?: number): string {
  return `<c r="${ref}"${styleAttribute(cellStyle)}><f>${escapeXml(formula)}</f></c>`;
}

function blankCell(ref: string, cellStyle: number): string {
  return `<c r="${ref}" s="${cellStyle}"/>`;
}

function row(index: number, cells: string[], height?: number): string {
  const heightAttributes =
    height === undefined ? '' : ` ht="${height}" customHeight="1"`;
  return `<row r="${index}"${heightAttributes}>${cells.join('')}</row>`;
}

function column(min: number, max: number, width: number): string {
  return `<col min="${min}" max="${max}" width="${width}" customWidth="1"/>`;
}

function mergeCells(refs: string[]): string {
  return `<mergeCells count="${refs.length}">${refs
    .map((ref) => `<mergeCell ref="${ref}"/>`)
    .join('')}</mergeCells>`;
}

function worksheet(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${body}</worksheet>`;
}

function relationships(entries: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
}

function relationship(id: string, type: string, target: string): string {
  return `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
}

const dashboardRows = [
  { region: 'North', units: 1240, revenue: 386400, growth: 0.182 },
  { region: 'South', units: 980, revenue: 274850, growth: 0.064 },
  { region: 'East', units: 1510, revenue: 512300, growth: 0.281 },
  { region: 'West', units: 1125, revenue: 331200, growth: -0.043 },
] as const;

const updatedSerial = 46203;
const dashboardFirstRow = 5;
const dashboardTotalRow = dashboardFirstRow + dashboardRows.length;
const currencyFormat = '"$"#,##0.00';

const dashboard = worksheet(
  [
    '<dimension ref="A1:F11"/>',
    `<cols>${[
      column(1, 1, 16),
      column(2, 2, 10),
      column(3, 4, 14),
      column(5, 5, 10),
      column(6, 6, 12),
    ].join('')}</cols>`,
    '<sheetData>',
    row(1, [textCell('A1', 'Q3 Sales Dashboard', style.title)], 30),
    row(2, [textCell('A2', 'Fiscal year 2026 — figures in USD', style.subtitle)]),
    row(4, [
      textCell('A4', 'Region', style.header),
      textCell('B4', 'Units', style.header),
      textCell('C4', 'Revenue', style.header),
      textCell('D4', 'Avg price', style.header),
      textCell('E4', 'Growth', style.header),
      textCell('F4', 'Updated', style.header),
    ]),
    ...dashboardRows.map((entry, index) => {
      const at = dashboardFirstRow + index;
      return row(at, [
        textCell(`A${at}`, entry.region),
        numberCell(`B${at}`, entry.units),
        numberCell(`C${at}`, entry.revenue, style.currency),
        formulaCell(`D${at}`, `C${at}/B${at}`, style.currency),
        numberCell(`E${at}`, entry.growth, style.percent),
        numberCell(`F${at}`, updatedSerial, style.date),
      ]);
    }),
    row(dashboardTotalRow, [
      textCell(`A${dashboardTotalRow}`, 'Total', style.totalLabel),
      formulaCell(
        `B${dashboardTotalRow}`,
        `SUM(B${dashboardFirstRow}:B${dashboardTotalRow - 1})`,
        style.totalPlain,
      ),
      formulaCell(
        `C${dashboardTotalRow}`,
        `SUM(C${dashboardFirstRow}:C${dashboardTotalRow - 1})`,
        style.totalCurrency,
      ),
      formulaCell(
        `D${dashboardTotalRow}`,
        `C${dashboardTotalRow}/B${dashboardTotalRow}`,
        style.totalCurrency,
      ),
      blankCell(`E${dashboardTotalRow}`, style.totalPlain),
      blankCell(`F${dashboardTotalRow}`, style.totalPlain),
    ]),
    row(11, [
      textCell('A11', 'Top performer: East — 28% growth on 1,510 units', style.callout),
      blankCell('B11', style.callout),
      blankCell('C11', style.callout),
      blankCell('D11', style.callout),
      blankCell('E11', style.callout),
      blankCell('F11', style.callout),
    ]),
    '</sheetData>',
    mergeCells(['A1:F1', 'A2:F2', 'A11:F11']),
    '<drawing r:id="rId1"/>',
  ].join(''),
);

const formulas = worksheet(
  [
    '<dimension ref="A1:K14"/>',
    `<cols>${[
      column(1, 1, 20),
      column(2, 2, 20),
      column(5, 5, 10),
      column(8, 9, 12),
      column(11, 11, 12),
    ].join('')}</cols>`,
    '<sheetData>',
    row(1, [textCell('A1', 'Formula playground', style.title)], 30),
    row(2, [
      textCell('A2', 'Function', style.header),
      textCell('B2', 'Result', style.header),
      textCell('E2', 'Values', style.header),
      textCell('H2', 'Key', style.header),
      textCell('I2', 'Value', style.header),
      textCell('K2', 'Region', style.header),
    ]),
    row(3, [
      textCell('A3', 'SUM', style.label),
      formulaCell('B3', 'SUM(E3:E7)'),
      numberCell('E3', 12),
      textCell('H3', 'Alpha'),
      numberCell('I3', 100),
      textCell('K3', 'North'),
    ]),
    row(4, [
      textCell('A4', 'AVERAGE', style.label),
      formulaCell('B4', 'AVERAGE(E3:E7)'),
      numberCell('E4', 47),
      textCell('H4', 'Beta'),
      numberCell('I4', 200),
      textCell('K4', 'South'),
    ]),
    row(5, [
      textCell('A5', 'MIN', style.label),
      formulaCell('B5', 'MIN(E3:E7)'),
      numberCell('E5', 8),
      textCell('H5', 'Gamma'),
      numberCell('I5', 300),
      textCell('K5', 'Northeast'),
    ]),
    row(6, [
      textCell('A6', 'MAX', style.label),
      formulaCell('B6', 'MAX(E3:E7)'),
      numberCell('E6', 33),
      textCell('K6', 'West'),
    ]),
    row(7, [
      textCell('A7', 'IF', style.label),
      formulaCell('B7', 'IF(B3>100,"Over budget","Within budget")'),
      numberCell('E7', 21),
    ]),
    row(8, [
      textCell('A8', 'IFS', style.label),
      formulaCell('B8', 'IFS(B4>=40,"High",B4>=20,"Medium",TRUE,"Low")'),
    ]),
    row(9, [
      textCell('A9', 'VLOOKUP', style.label),
      formulaCell('B9', 'VLOOKUP("Gamma",H3:I5,2,FALSE)'),
    ]),
    row(10, [
      textCell('A10', 'COUNTIF (wildcard)', style.label),
      formulaCell('B10', 'COUNTIF(K3:K6,"North*")'),
    ]),
    row(11, [
      textCell('A11', 'Percent of max', style.label),
      formulaCell('B11', 'B4/B6', style.percent),
    ]),
    row(12, [
      textCell('A12', 'TEXT', style.label),
      formulaCell('B12', 'TEXT(B4,"#,##0.00")'),
    ]),
    row(13, [
      textCell('A13', 'TODAY (volatile)', style.label),
      formulaCell('B13', 'TODAY()', style.date),
    ]),
    row(14, [
      textCell('A14', 'End of month', style.label),
      formulaCell('B14', 'EOMONTH(TODAY(),0)', style.date),
    ]),
    '</sheetData>',
    mergeCells(['A1:B1']),
  ].join(''),
);

/** `region|product|rep|units|unitPrice|margin`, one line per transaction. */
const transactions = `
South|Bracket|Blair|201|9.86|0.401
North|Flange|Finley|80|89.19|0.164
South|Gadget|Casey|124|32.06|0.159
East|Cog|Harper|57|84.06|0.223
West|Flange|Harper|26|97.3|0.222
North|Cog|Finley|146|18.98|0.348
South|Sprocket|Devon|93|73.34|0.385
West|Cog|Devon|58|12.9|0.429
South|Cog|Casey|29|33.21|0.218
West|Gadget|Emery|187|67.34|0.425
North|Bracket|Emery|100|14.46|0.315
South|Sprocket|Avery|13|22.04|0.407
South|Flange|Devon|34|77.93|0.297
North|Widget|Casey|122|78.33|0.128
South|Sprocket|Devon|85|73.83|0.091
West|Bracket|Harper|79|67.9|0.273
North|Bracket|Avery|73|87.37|0.108
West|Widget|Harper|107|49.76|0.31
East|Bracket|Casey|107|78|0.382
East|Gadget|Blair|36|52.5|0.121
West|Gadget|Casey|54|19.32|0.186
North|Bracket|Blair|137|12.8|0.191
South|Sprocket|Devon|108|38.73|0.449
South|Cog|Devon|108|80.09|0.239
South|Gadget|Blair|9|14.63|0.364
East|Flange|Devon|7|60.03|0.165
West|Flange|Harper|156|35.37|0.263
North|Widget|Emery|36|55.94|0.141
West|Sprocket|Blair|41|40.63|0.345
North|Sprocket|Harper|169|49.95|0.147
South|Sprocket|Devon|126|55.79|0.166
East|Widget|Avery|78|42.42|0.295
South|Widget|Devon|121|20.24|0.331
South|Sprocket|Emery|39|27.93|0.376
West|Sprocket|Blair|34|53.02|0.33
South|Flange|Emery|78|58.47|0.389
North|Cog|Casey|18|24.26|0.378
North|Bracket|Finley|172|26.01|0.097
South|Cog|Harper|178|90.65|0.257
South|Bracket|Emery|170|9.84|0.175
South|Bracket|Blair|135|65.44|0.18
North|Bracket|Casey|29|47.02|0.409
South|Gadget|Finley|123|34.4|0.207
North|Flange|Devon|201|46.33|0.238
West|Sprocket|Harper|82|31.7|0.303
South|Gadget|Devon|201|21.65|0.357
North|Flange|Harper|159|75.33|0.219
West|Cog|Blair|184|11.47|0.159
West|Flange|Casey|110|72.06|0.459
East|Gadget|Blair|175|68.99|0.136
North|Cog|Harper|45|32.83|0.353
North|Cog|Devon|47|62.09|0.164
North|Sprocket|Blair|13|62.93|0.47
East|Cog|Emery|108|84.46|0.352
North|Bracket|Devon|17|52.29|0.107
West|Bracket|Casey|15|75.37|0.369
West|Widget|Casey|61|45.43|0.129
West|Flange|Blair|105|27.49|0.309
North|Bracket|Finley|118|41.89|0.242
East|Widget|Harper|82|58.54|0.436
South|Cog|Emery|13|62.13|0.098
East|Bracket|Devon|118|18.99|0.444
East|Cog|Harper|136|30.11|0.459
North|Gadget|Devon|111|11.27|0.169
East|Sprocket|Finley|85|44.52|0.123
North|Flange|Blair|57|25.71|0.355
North|Flange|Casey|26|25.95|0.474
North|Bracket|Finley|13|43.95|0.086
South|Gadget|Finley|167|21.45|0.13
West|Cog|Emery|204|89.13|0.244
North|Sprocket|Finley|161|39.58|0.39
South|Sprocket|Emery|151|23.45|0.38
South|Widget|Devon|85|78.28|0.144
North|Cog|Avery|73|39.68|0.224
East|Bracket|Emery|33|91.49|0.36
North|Flange|Finley|59|20.44|0.195
South|Sprocket|Devon|106|82.77|0.424
North|Widget|Casey|167|74.56|0.294
East|Sprocket|Avery|171|97.12|0.379
South|Gadget|Finley|24|37.43|0.326
West|Widget|Avery|165|74.75|0.218
West|Bracket|Avery|93|34.33|0.421
East|Cog|Devon|10|10.09|0.095
South|Gadget|Harper|36|54.2|0.403
East|Flange|Devon|80|44.42|0.275
South|Cog|Avery|200|85.48|0.124
West|Sprocket|Finley|68|14.31|0.298
North|Sprocket|Avery|5|57.74|0.149
West|Sprocket|Devon|91|13.74|0.267
West|Sprocket|Blair|41|63.71|0.289
West|Flange|Avery|85|29.83|0.095
South|Widget|Devon|48|94.41|0.454
West|Widget|Casey|115|91.82|0.272
South|Widget|Devon|38|34.31|0.402
West|Sprocket|Finley|129|57.31|0.282
South|Gadget|Casey|69|64.14|0.138
West|Bracket|Emery|109|53.11|0.181
North|Sprocket|Avery|56|41.89|0.451
South|Widget|Casey|23|81.78|0.4
South|Flange|Blair|32|70.28|0.376
West|Cog|Devon|36|58.8|0.427
West|Cog|Emery|54|51.58|0.284
East|Cog|Blair|109|76.55|0.326
West|Sprocket|Blair|99|32.92|0.273
North|Flange|Blair|198|16|0.328
North|Sprocket|Emery|149|79.04|0.367
West|Flange|Casey|7|13.36|0.474
West|Cog|Harper|114|77.17|0.47
North|Flange|Avery|30|52.02|0.463
East|Flange|Casey|189|69.84|0.251
North|Widget|Blair|36|76.35|0.358
North|Sprocket|Finley|150|70.42|0.34
West|Widget|Avery|158|72.71|0.182
South|Widget|Casey|201|41.07|0.256
South|Gadget|Blair|95|30.08|0.477
North|Bracket|Emery|166|38.26|0.319
East|Cog|Emery|81|31.3|0.098
South|Flange|Harper|188|53.12|0.349
West|Cog|Harper|20|41.47|0.152
East|Flange|Harper|179|8.46|0.479
West|Gadget|Casey|128|88.86|0.379
North|Flange|Finley|123|89.71|0.442
West|Cog|Harper|167|41.7|0.355
South|Sprocket|Devon|31|76.73|0.164
East|Cog|Finley|168|41.04|0.443
South|Widget|Casey|81|15.66|0.178
North|Bracket|Harper|191|47.55|0.202
West|Bracket|Finley|88|42.36|0.469
South|Sprocket|Emery|90|84.44|0.476
West|Flange|Blair|28|36.39|0.285
South|Sprocket|Avery|164|57.15|0.358
North|Flange|Blair|177|12.72|0.461
West|Bracket|Avery|21|46.33|0.449
East|Widget|Avery|60|59.01|0.457
East|Cog|Avery|91|54.65|0.403
West|Gadget|Harper|9|76.5|0.451
North|Gadget|Casey|8|12.98|0.347
West|Gadget|Blair|99|68.72|0.355
South|Bracket|Harper|154|12.54|0.234
South|Flange|Casey|74|90.49|0.422
North|Sprocket|Casey|102|59.84|0.108
West|Cog|Finley|48|43.42|0.182
North|Sprocket|Blair|180|50.64|0.448
North|Flange|Harper|120|49.73|0.414
East|Bracket|Avery|197|70.73|0.243
North|Bracket|Devon|58|74.95|0.339
South|Sprocket|Finley|105|43.36|0.43
West|Flange|Emery|42|94.19|0.466
South|Cog|Avery|56|16.81|0.448
West|Cog|Avery|97|86.58|0.375
East|Sprocket|Avery|94|19.47|0.137
North|Gadget|Blair|157|32.23|0.409
West|Gadget|Devon|114|21.31|0.409
North|Bracket|Casey|95|96.89|0.169
East|Bracket|Finley|156|63.82|0.228
North|Sprocket|Finley|75|64.69|0.19
East|Cog|Harper|195|31.24|0.322
North|Cog|Finley|187|75.61|0.407
East|Sprocket|Casey|126|99.45|0.398
East|Widget|Harper|107|69.94|0.18
West|Gadget|Devon|101|19.09|0.28
East|Widget|Devon|188|90.05|0.097
East|Widget|Harper|172|61.54|0.325
South|Cog|Harper|151|97.07|0.198
North|Flange|Devon|171|84.19|0.362
West|Gadget|Devon|59|15.46|0.223
South|Cog|Harper|144|34.57|0.346
South|Bracket|Finley|98|58.34|0.105
West|Bracket|Emery|106|40.01|0.21
South|Gadget|Casey|142|54.86|0.438
West|Sprocket|Emery|46|72.08|0.406
North|Widget|Emery|110|33.32|0.199
South|Sprocket|Emery|37|64.06|0.306
West|Widget|Finley|56|93.33|0.373
North|Bracket|Emery|24|64.08|0.126
North|Cog|Harper|137|74.18|0.403
South|Bracket|Devon|197|65.36|0.187
East|Sprocket|Blair|162|26.36|0.277
North|Gadget|Blair|139|63.58|0.21
West|Bracket|Casey|174|41.8|0.115
East|Widget|Avery|89|51.17|0.456
East|Gadget|Harper|67|55.21|0.137
West|Gadget|Finley|29|83.18|0.395
South|Widget|Devon|154|61.19|0.354
South|Widget|Blair|69|37.05|0.407
North|Gadget|Casey|134|45.56|0.473
East|Bracket|Avery|143|29.92|0.43
West|Bracket|Emery|47|85.98|0.164
South|Sprocket|Avery|192|27.8|0.352
South|Widget|Harper|194|52.39|0.415
South|Flange|Emery|124|89.89|0.167
West|Widget|Avery|89|56.2|0.447
North|Widget|Casey|81|98.01|0.457
North|Widget|Finley|165|74.88|0.421
North|Gadget|Finley|8|50.47|0.406
South|Cog|Harper|127|67.14|0.195
East|Gadget|Finley|158|76.12|0.369
East|Sprocket|Avery|60|64.34|0.252
South|Bracket|Devon|79|95.76|0.321
South|Cog|Blair|45|29.36|0.308
West|Gadget|Avery|167|25.42|0.376
South|Sprocket|Emery|7|95.9|0.344
North|Sprocket|Devon|38|85.46|0.46
West|Cog|Blair|73|40.53|0.179
West|Flange|Devon|40|92.63|0.443
West|Sprocket|Harper|177|34.53|0.234
West|Cog|Avery|199|18.36|0.472
East|Gadget|Finley|127|66.97|0.251
North|Sprocket|Emery|132|74.76|0.15
East|Flange|Devon|47|67.53|0.47
West|Bracket|Devon|165|30.6|0.429
East|Sprocket|Emery|38|16.38|0.127
South|Gadget|Finley|186|29.34|0.289
North|Cog|Blair|54|17.35|0.087
North|Cog|Finley|182|96.25|0.436
South|Widget|Casey|82|37.39|0.187
North|Flange|Finley|163|9.75|0.373
South|Sprocket|Emery|160|39.21|0.464
South|Flange|Devon|26|84.02|0.271
North|Flange|Emery|190|67.41|0.119
West|Sprocket|Emery|21|42.29|0.268
East|Sprocket|Devon|83|92.42|0.262
West|Gadget|Casey|163|22.62|0.22
South|Gadget|Finley|179|55.58|0.235
West|Bracket|Blair|32|98.1|0.095
North|Cog|Harper|165|16.9|0.216
East|Flange|Harper|181|79.11|0.119
North|Widget|Devon|124|84.91|0.447
South|Cog|Harper|116|88.39|0.421
East|Gadget|Harper|128|82.87|0.285
South|Cog|Devon|60|63.06|0.209
South|Gadget|Finley|24|29.22|0.193
East|Widget|Emery|86|83.76|0.296
North|Flange|Harper|102|31.23|0.152
North|Flange|Harper|75|18.79|0.413
South|Widget|Casey|155|61.7|0.474
West|Gadget|Avery|166|24.59|0.294
South|Widget|Finley|13|29.31|0.317
East|Sprocket|Avery|16|51.28|0.331
South|Flange|Finley|156|76.15|0.146
West|Widget|Emery|184|25.28|0.164
South|Widget|Finley|116|31.32|0.081
East|Flange|Emery|121|9.49|0.311
South|Bracket|Avery|71|87.11|0.229
East|Sprocket|Avery|53|42.98|0.367
East|Sprocket|Finley|86|71.31|0.414
North|Widget|Emery|63|65.22|0.134
West|Sprocket|Emery|61|40.36|0.362
West|Cog|Emery|13|61.96|0.111
North|Bracket|Finley|121|73.81|0.314
North|Bracket|Devon|21|44.15|0.177
East|Gadget|Blair|68|93.09|0.475
South|Bracket|Avery|81|31.8|0.205
North|Gadget|Avery|8|17.58|0.324
East|Widget|Finley|61|71.3|0.23
West|Gadget|Avery|139|43.11|0.421
North|Cog|Finley|30|55.08|0.371
West|Widget|Finley|59|30.23|0.402
North|Cog|Casey|61|32.24|0.276
South|Flange|Harper|29|15.45|0.202
East|Gadget|Avery|83|63.34|0.458
South|Flange|Emery|51|99.75|0.108
East|Sprocket|Finley|113|71.67|0.185
West|Sprocket|Devon|193|23.19|0.249
South|Flange|Devon|176|91.66|0.398
East|Gadget|Blair|142|97.51|0.176
West|Gadget|Casey|49|14.16|0.109
West|Flange|Emery|19|30.76|0.47
East|Flange|Casey|120|78.19|0.267
West|Cog|Avery|57|87.6|0.471
North|Flange|Avery|77|98.41|0.204
North|Widget|Blair|83|75|0.39
West|Cog|Harper|101|26.54|0.312
West|Bracket|Devon|73|84.57|0.261
East|Gadget|Avery|151|85.14|0.168
West|Sprocket|Devon|195|47.26|0.454
West|Bracket|Harper|28|61.95|0.266
North|Flange|Emery|133|17|0.13
East|Widget|Devon|139|93.47|0.453
West|Bracket|Finley|43|30.69|0.197
South|Flange|Devon|126|30.94|0.22
East|Flange|Avery|73|80.5|0.472
North|Gadget|Avery|137|78.66|0.44
North|Widget|Finley|131|23.37|0.466
West|Cog|Blair|17|47.61|0.222
East|Widget|Casey|45|50.64|0.308
North|Gadget|Casey|123|96.32|0.36
West|Gadget|Finley|82|20.79|0.115
West|Bracket|Harper|20|60.78|0.327
East|Bracket|Casey|63|57.14|0.36
South|Gadget|Harper|102|96.04|0.369
South|Bracket|Casey|84|54.12|0.294
West|Flange|Finley|99|30|0.214
West|Bracket|Casey|177|68.5|0.349
South|Flange|Blair|30|13.94|0.323
East|Bracket|Blair|92|85.07|0.409
East|Sprocket|Casey|125|78.55|0.296
North|Widget|Avery|16|33.35|0.087
East|Gadget|Finley|28|77.03|0.364
North|Cog|Finley|22|70.55|0.387
`
  .trim()
  .split('\n')
  .map((line) => {
    const [region, product, rep, units, unitPrice, margin] = line.split('|');
    return {
      region: region!,
      product: product!,
      rep: rep!,
      units: Number(units),
      unitPrice: Number(unitPrice),
      margin: Number(margin),
    };
  });

const firstTransactionSerial = 45658;

const data = worksheet(
  [
    `<dimension ref="A1:H${transactions.length + 1}"/>`,
    `<cols>${[
      column(1, 1, 12),
      column(2, 4, 12),
      column(5, 5, 9),
      column(6, 7, 13),
      column(8, 8, 10),
    ].join('')}</cols>`,
    '<sheetData>',
    row(1, [
      textCell('A1', 'Date', style.header),
      textCell('B1', 'Region', style.header),
      textCell('C1', 'Product', style.header),
      textCell('D1', 'Rep', style.header),
      textCell('E1', 'Units', style.header),
      textCell('F1', 'Unit price', style.header),
      textCell('G1', 'Revenue', style.header),
      textCell('H1', 'Margin', style.header),
    ]),
    ...transactions.map((entry, index) => {
      const at = index + 2;
      return row(at, [
        numberCell(`A${at}`, firstTransactionSerial + index, style.date),
        textCell(`B${at}`, entry.region),
        textCell(`C${at}`, entry.product),
        textCell(`D${at}`, entry.rep),
        numberCell(`E${at}`, entry.units),
        numberCell(`F${at}`, entry.unitPrice, style.currency),
        formulaCell(`G${at}`, `E${at}*F${at}`, style.currency),
        numberCell(`H${at}`, entry.margin, style.percent),
      ]);
    }),
    '</sheetData>',
  ].join(''),
);

function cachePoints(values: (string | number)[]): string {
  return `<c:ptCount val="${values.length}"/>${values
    .map(
      (value, index) =>
        `<c:pt idx="${index}"><c:v>${escapeXml(String(value))}</c:v></c:pt>`,
    )
    .join('')}`;
}

const categoryAxis = 553771648;
const valueAxis = 553773184;

const chart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Revenue by region</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Dashboard!$C$4</c:f><c:strCache>${cachePoints(['Revenue'])}</c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>
          <c:invertIfNegative val="0"/>
          <c:cat><c:strRef><c:f>Dashboard!$A$${dashboardFirstRow}:$A$${dashboardTotalRow - 1}</c:f><c:strCache>${cachePoints(
            dashboardRows.map((entry) => entry.region),
          )}</c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Dashboard!$C$${dashboardFirstRow}:$C$${dashboardTotalRow - 1}</c:f><c:numCache><c:formatCode>${escapeXml(
            currencyFormat,
          )}</c:formatCode>${cachePoints(
            dashboardRows.map((entry) => entry.revenue),
          )}</c:numCache></c:numRef></c:val>
        </c:ser>
        <c:gapWidth val="150"/>
        <c:axId val="${categoryAxis}"/>
        <c:axId val="${valueAxis}"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="${categoryAxis}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="${valueAxis}"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
        <c:noMultiLvlLbl val="0"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="${valueAxis}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="${escapeXml('"$"#,##0')}" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="${categoryAxis}"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>`;

const drawing = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>14</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Revenue by region"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="${escapeXml(
  currencyFormat,
)}"/></numFmts><fonts count="5"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><sz val="10"/><color rgb="FF808080"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.5999938962981048"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top style="thin"><color rgb="FF000000"/></top><bottom style="double"><color rgb="FF000000"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="12"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/><xf numFmtId="164" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>`;

const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="0"/><sheets><sheet name="Dashboard" sheetId="1" r:id="rId1"/><sheet name="Formulas" sheetId="2" r:id="rId2"/><sheet name="Data" sheetId="3" r:id="rId3"/></sheets></workbook>`;

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`;

const rootRelationships = relationships(
  relationship('rId1', 'officeDocument', 'xl/workbook.xml'),
);

const workbookRelationships = relationships(
  [
    relationship('rId1', 'worksheet', 'worksheets/sheet1.xml'),
    relationship('rId2', 'worksheet', 'worksheets/sheet2.xml'),
    relationship('rId3', 'worksheet', 'worksheets/sheet3.xml'),
    relationship('rId4', 'sharedStrings', 'sharedStrings.xml'),
    relationship('rId5', 'styles', 'styles.xml'),
  ].join(''),
);

const dashboardRelationships = relationships(
  relationship('rId1', 'drawing', '../drawings/drawing1.xml'),
);

const drawingRelationships = relationships(
  relationship('rId1', 'chart', '../charts/chart1.xml'),
);

const sharedStringTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStringUses}" uniqueCount="${
  sharedStrings.length
}">${sharedStrings
  .map((value) => `<si><t xml:space="preserve">${escapeXml(value)}</t></si>`)
  .join('')}</sst>`;

const archive = new JSZip();
const options = { date: zipDate, createFolders: false };
const textParts: Record<string, string> = {
  '[Content_Types].xml': contentTypes,
  '_rels/.rels': rootRelationships,
  'xl/workbook.xml': workbook,
  'xl/_rels/workbook.xml.rels': workbookRelationships,
  'xl/worksheets/sheet1.xml': dashboard,
  'xl/worksheets/_rels/sheet1.xml.rels': dashboardRelationships,
  'xl/worksheets/sheet2.xml': formulas,
  'xl/worksheets/sheet3.xml': data,
  'xl/drawings/drawing1.xml': drawing,
  'xl/drawings/_rels/drawing1.xml.rels': drawingRelationships,
  'xl/charts/chart1.xml': chart,
  'xl/styles.xml': styles,
  'xl/sharedStrings.xml': sharedStringTable,
};

for (const [partPath, contents] of Object.entries(textParts)) {
  archive.file(partPath, contents, options);
}

const buffer = await archive.generateAsync({
  type: 'nodebuffer',
  compression: 'DEFLATE',
  compressionOptions: { level: 9 },
});
fs.mkdirSync(path.dirname(output), { recursive: true });
fs.writeFileSync(output, buffer);
console.log(`Created ${output} (${buffer.length} bytes)`);
