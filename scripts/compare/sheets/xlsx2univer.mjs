// Neutral xlsx -> Univer IWorkbookData converter (Univer's OSS build imports no xlsx).
// Maps values, formulas, number formats, fonts, fills, borders, alignment, merges, column widths, row heights,
// frozen panes, hidden rows/columns/sheets, gridline visibility, defined names, Excel tables (sheets-table plugin resource)
// and sheet order. Sheet extents cover every explicit reference in formulas and defined names, up to Excel's limits.
// Counts what it cannot map (CF, data validation, charts, images, shapes, comments, sparklines, pivots,
// hyperlinks, rich-text runs, gradient/pattern fills, indents) in `report.unmapped`; those are "not loadable in Univer OSS".
// Column widths are written in px from a fixed max-digit-width table and also kept in characters under `custom`,
// so a page that measures the real font can recompute them. Usage: node xlsx2univer.mjs in.xlsx out.json [report.json]
import { createRequire } from 'node:module';
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';

const require = createRequire(resolve(import.meta.dirname, 'site/package.json'));
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const BORDER = { thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8, mediumDashed: 9,
  mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13 };
const HALIGN = { left: 1, center: 2, right: 3, justify: 4, distributed: 6, fill: 1, centerContinuous: 2 };
const VALIGN = { top: 1, middle: 2, center: 2, bottom: 3, justify: 1, distributed: 2 };
const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00',
  '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF',
  '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000',
  '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00',
  'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
const MDW = { calibri: 7, carlito: 7, arial: 7, liberation: 7, helvetica: 7, cambria: 7, 'times new roman': 7, verdana: 8, tahoma: 7, 'segoe ui': 7, aptos: 7 };
const ERRORS = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)$/;

export function normalizeFormula(text) {
  return text
    .replace(/_xlfn\.ANCHORARRAY\(\s*((?:'[^']+'|[A-Za-z0-9_.]+)?!?\$?[A-Z]{1,3}\$?\d+)\s*\)/g, '$1#')
    .replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\.|_xlpm\./g, '');
}

function xmlText(value) {
  return value.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

function themeColors(xml) {
  const scheme = /<a:clrScheme[\s\S]*?<\/a:clrScheme>/.exec(xml ?? '')?.[0] ?? '';
  const pick = (tag) => {
    const block = new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`).exec(scheme)?.[1] ?? '';
    return /lastClr="([0-9A-Fa-f]{6})"/.exec(block)?.[1] ?? /srgbClr val="([0-9A-Fa-f]{6})"/.exec(block)?.[1] ?? null;
  };
  const [dk1, lt1, dk2, lt2] = ['dk1', 'lt1', 'dk2', 'lt2'].map(pick);
  return [lt1 ?? 'FFFFFF', dk1 ?? '000000', lt2 ?? 'E7E6E6', dk2 ?? '44546A',
    ...['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map((t) => pick(t) ?? '000000')];
}

function tint(hex, amount) {
  if (!amount) return hex;
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = amount < 0 ? l * (1 + amount) : l * (1 - amount) + amount;
  const hue = (p, q, t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3); }
  return [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function serial(date, date1904) {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (date.getTime() - epoch) / 86_400_000;
}

const colIndex = (letters) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
function rangeOf(a1) {
  const [a, b = a] = a1.replace(/\$/g, '').split(':');
  const p = (s) => { const m = /^([A-Z]{1,3})(\d+)$/.exec(s); return m ? { r: Number(m[2]) - 1, c: colIndex(m[1]) } : null; };
  const s = p(a), e = p(b);
  return s && e ? { startRow: Math.min(s.r, e.r), endRow: Math.max(s.r, e.r), startColumn: Math.min(s.c, e.c), endColumn: Math.max(s.c, e.c) } : null;
}

async function packageFacts(zip) {
  const facts = { cfRules: 0, cfRanges: 0, validations: 0, charts: 0, images: 0, shapes: 0, comments: 0, threadedComments: 0,
    sparklines: 0, pivots: 0, tables: 0, sheetHyperlinks: 0, chartsheets: 0, x14cf: 0 };
  for (const name of Object.keys(zip.files)) {
    if (/^xl\/charts\/chart\d*\.xml$/.test(name)) facts.charts++;
    else if (/^xl\/media\//.test(name)) facts.images++;
    else if (/^xl\/comments\d*\.xml$/.test(name)) facts.comments += ((await zip.file(name).async('string')).match(/<comment /g) ?? []).length;
    else if (/^xl\/threadedComments\//.test(name)) facts.threadedComments++;
    else if (/^xl\/pivotTables\//.test(name) && name.endsWith('.xml')) facts.pivots++;
    else if (/^xl\/tables\/table\d*\.xml$/.test(name)) facts.tables++;
    else if (/^xl\/chartsheets\/sheet\d*\.xml$/.test(name)) facts.chartsheets++;
    else if (/^xl\/drawings\/drawing\d*\.xml$/.test(name)) facts.shapes += ((await zip.file(name).async('string')).match(/<xdr:sp[ >]/g) ?? []).length;
    else if (/^xl\/worksheets\/sheet\d*\.xml$/.test(name)) {
      const xml = await zip.file(name).async('string');
      facts.cfRanges += (xml.match(/<conditionalFormatting[ >]/g) ?? []).length;
      facts.cfRules += (xml.match(/<cfRule[ >]/g) ?? []).length;
      facts.x14cf += (xml.match(/<x14:cfRule[ >]/g) ?? []).length;
      facts.validations += (xml.match(/<dataValidation[ >]/g) ?? []).length + (xml.match(/<x14:dataValidation[ >]/g) ?? []).length;
      facts.sparklines += (xml.match(/<x14:sparkline>/g) ?? []).length;
      facts.sheetHyperlinks += (xml.match(/<hyperlink /g) ?? []).length;
    }
  }
  return facts;
}

const MAX_ROWS = 1_048_576, MAX_COLS = 16_384;
const REF = /(?<![\w.$\u00C0-\uFFFF'])(?:'((?:[^']|'')+)'!|([A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*)!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![\w(!.\u00C0-\uFFFF])/g;

/** Calls visit(sheetName | null, row, col) for the far corner of every explicit A1 reference in a formula. */
export function scanRefs(text, visit) {
  let body = text.replace(/"(?:[^"]|"")*"/g, '""');
  for (let prev = ''; prev !== body;) { prev = body; body = body.replace(/\[[^\[\]]*\]/g, ''); }
  for (const m of body.matchAll(REF)) {
    const sheet = m[1] != null ? m[1].replace(/''/g, "'") : m[2] ?? null;
    const row = Math.max(Number(m[4]), Number(m[6] ?? 0)) - 1;
    const col = Math.max(colIndex(m[3]), m[5] ? colIndex(m[5]) : 0);
    if (row < MAX_ROWS && col < MAX_COLS) visit(sheet, row, col);
  }
}

function workbookSheets(workbookXml) {
  return [...workbookXml.matchAll(/<sheet\b([^>]*?)\/?>/g)].map((m) => {
    const attrs = Object.fromEntries([...m[1].matchAll(/([\w:]+)="([^"]*)"/g)].map((a) => [a[1], xmlText(a[2])]));
    return { name: attrs.name, rid: attrs['r:id'] };
  });
}

function definedNames(workbookXml, scopeOf) {
  const out = {};
  let n = 0;
  for (const m of workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const attrs = Object.fromEntries([...m[1].matchAll(/(\w+)="([^"]*)"/g)].map((a) => [a[1], xmlText(a[2])]));
    if (!attrs.name || /^_xlnm\.|^_xlfn\.|^_xlpm\./i.test(attrs.name)) continue;
    const body = normalizeFormula(xmlText(m[2]).trim());
    const isRef = /^(?:'[^']+'|[^'!\s(),]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/.test(body);
    const id = `dn${n++}`;
    out[id] = { id, name: attrs.name, formulaOrRefString: isRef ? body : `=${body}`, hidden: attrs.hidden === '1',
      localSheetId: attrs.localSheetId != null ? scopeOf(Number(attrs.localSheetId)) ?? 'AllDefaultWorkbook' : 'AllDefaultWorkbook' };
  }
  return out;
}

const relTargets = (relsXml, type) => [...(relsXml ?? '').matchAll(/<Relationship\b([^>]*)\/?>/g)]
  .map((m) => Object.fromEntries([...m[1].matchAll(/(\w+)="([^"]*)"/g)].map((a) => [a[1], a[2]])))
  .filter((r) => !type || new RegExp(`/${type}$`).test(r.Type ?? ''));
const joinPart = (base, target) => {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
  return parts.join('/');
};

/** Excel tables per worksheet name, in the sheets-table plugin's snapshot shape (header/totals rows, column titles). */
async function excelTables(zip, sheetsInBook) {
  const wbRels = relTargets(await zip.file('xl/_rels/workbook.xml.rels')?.async('string'));
  const out = new Map();
  let n = 0;
  for (const sheet of sheetsInBook) {
    const rel = wbRels.find((r) => r.Id === sheet.rid);
    if (!rel) continue;
    const part = joinPart('xl/workbook.xml', rel.Target);
    const rels = await zip.file(part.replace(/([^/]+)$/, '_rels/$1.rels'))?.async('string');
    for (const t of relTargets(rels, 'table')) {
      const xml = await zip.file(joinPart(part, t.Target))?.async('string');
      if (!xml) continue;
      const head = /<table\b([^>]*)>/.exec(xml)?.[1] ?? '';
      const attr = (k) => { const v = new RegExp(`\\b${k}="([^"]*)"`).exec(head)?.[1]; return v == null ? null : xmlText(v); };
      const range = rangeOf(attr('ref') ?? '');
      const name = attr('displayName') ?? attr('name');
      if (!range || !name) continue;
      const columns = [...xml.matchAll(/<tableColumn\b([^>]*?)\/?>/g)].map((c, i) => ({
        id: `tb${n}c${i}`, displayName: xmlText(/\bname="([^"]*)"/.exec(c[1])?.[1] ?? `Column${i + 1}`)
          .replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16))) }));
      const showHeader = attr('headerRowCount') !== '0';
      const showFooter = Number(attr('totalsRowCount') ?? 0) > 0;
      (out.get(sheet.name) ?? out.set(sheet.name, []).get(sheet.name)).push({
        id: `tb${n++}`, name, range, options: { showHeader, showFooter, showAutoFilter: false }, columns, meta: {} });
    }
  }
  return out;
}

export async function convert(bytes, { id = 'wb', keepFormulaValues = true } = {}) {
  const zip = await JSZip.loadAsync(bytes);
  const facts = await packageFacts(zip);
  const theme = themeColors(await zip.file('xl/theme/theme1.xml')?.async('string'));
  const workbookXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? '';
  const date1904 = /<workbookPr[^>]*date1904="(1|true)"/.test(workbookXml);
  const stylesXml = (await zip.file('xl/styles.xml')?.async('string')) ?? '';
  const font0 = /<fonts\b[^>]*>\s*<font\b[^>]*>([\s\S]*?)<\/font>/.exec(stylesXml)?.[1] ?? '';
  const defaultFont = { ff: xmlText(/<name val="([^"]*)"/.exec(font0)?.[1] ?? 'Calibri'), fs: Number(/<sz val="([^"]*)"/.exec(font0)?.[1] ?? 11) };
  const sheetsInBook = workbookSheets(workbookXml);
  const realName = new Map();
  const taken = new Set(sheetsInBook.map((sh) => sh.name));
  zip.file('xl/workbook.xml', workbookXml.replace(/(<sheet\b[^>]*?\bname=")History(")/g, (_, a, b) => {
    let alias = 'History_';
    while (taken.has(alias)) alias += '_';
    taken.add(alias);
    realName.set(alias, 'History');
    return `${a}${alias}${b}`;
  }));
  const tables = await excelTables(zip, sheetsInBook);
  for (const name of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/[^/]+\.xml$/.test(n))) {
    const xml = await zip.file(name).async('string');
    zip.file(name, xml.replace(/<(?:\w+:)?(?:drawing|legacyDrawing|legacyDrawingHF)\b[^>]*\/>/g, '').replace(/<(?:\w+:)?tableParts\b[^>]*\/>/g, '').replace(/<(?:\w+:)?tableParts\b[^>]*>[\s\S]*?<\/(?:\w+:)?tableParts>/g, ''));
  }
  for (const name of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/_rels\/.*\.rels$/.test(n))) {
    const xml = await zip.file(name).async('string');
    zip.file(name, xml.replace(/<Relationship\b[^>]*Type="[^"]*\/(?:drawing|vmlDrawing|comments|table|pivotTable|threadedComment)"[^>]*\/>/g, ''));
  }
  for (const name of Object.keys(zip.files).filter((n) => /^xl\/(?:drawings|tables|pivotTables)\//.test(n))) zip.remove(name);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await zip.generateAsync({ type: 'nodebuffer' }));
  const unmapped = { conditionalFormatting: facts.cfRules + facts.x14cf, dataValidation: facts.validations, charts: facts.charts,
    images: facts.images, shapes: facts.shapes, comments: facts.comments + facts.threadedComments, sparklines: facts.sparklines,
    pivotTables: facts.pivots, tables: facts.tables, chartsheets: facts.chartsheets, hyperlinks: facts.sheetHyperlinks,
    richTextCells: 0, gradientFills: 0, patternFills: 0, indents: 0, columnStyles: 0, rowStyles: 0 };
  const color = (c, fallback = null) => {
    if (!c) return fallback;
    if (c.argb && /^[0-9A-Fa-f]{8}$/.test(c.argb)) return `#${c.argb.slice(2).toUpperCase()}`;
    if (c.theme != null) return `#${tint(theme[c.theme] ?? '000000', c.tint ?? 0)}`;
    if (c.indexed != null) return c.indexed === 64 ? fallback : `#${INDEXED[c.indexed] ?? '000000'}`;
    return fallback;
  };
  const styles = {};
  const styleIds = new Map();
  const styleId = (style) => {
    if (style.ff === defaultFont.ff) delete style.ff;
    if (style.fs === defaultFont.fs) delete style.fs;
    const key = JSON.stringify(style);
    if (key === '{}') return null;
    if (!styleIds.has(key)) { const sid = `s${styleIds.size}`; styleIds.set(key, sid); styles[sid] = style; }
    return styleIds.get(key);
  };
  const borderSide = (side) => (side?.style && BORDER[side.style] ? { s: BORDER[side.style], cl: { rgb: color(side.color, '#000000') } } : null);
  const mapStyle = (cell) => {
    const out = {};
    const f = cell.font;
    if (f) {
      if (f.name) out.ff = f.name;
      if (f.size) out.fs = f.size;
      if (f.bold) out.bl = 1;
      if (f.italic) out.it = 1;
      if (f.underline && f.underline !== 'none') out.ul = { s: 1 };
      if (f.strike) out.st = { s: 1 };
      const fc = color(f.color);
      if (fc && fc !== '#000000') out.cl = { rgb: fc };
      if (f.vertAlign === 'superscript') out.va = 3;
      if (f.vertAlign === 'subscript') out.va = 2;
    }
    const fill = cell.fill;
    if (fill?.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
      const bg = color(fill.fgColor) ?? color(fill.bgColor);
      if (fill.pattern !== 'solid') unmapped.patternFills++;
      if (bg) out.bg = { rgb: bg };
    } else if (fill?.type === 'gradient') {
      unmapped.gradientFills++;
      const stop = fill.stops?.[0]?.color;
      if (stop) out.bg = { rgb: color(stop, '#FFFFFF') };
    }
    const b = cell.border;
    if (b) {
      const bd = {};
      for (const [k, side] of [['t', b.top], ['r', b.right], ['b', b.bottom], ['l', b.left]]) { const v = borderSide(side); if (v) bd[k] = v; }
      if (b.diagonal?.style) {
        const v = borderSide(b.diagonal);
        if (v && b.diagonal.down) bd.tl_br = v;
        if (v && b.diagonal.up) bd.bl_tr = v;
      }
      if (Object.keys(bd).length) out.bd = bd;
    }
    const a = cell.alignment;
    if (a) {
      if (HALIGN[a.horizontal]) out.ht = HALIGN[a.horizontal];
      if (VALIGN[a.vertical]) out.vt = VALIGN[a.vertical];
      if (a.wrapText) out.tb = 3;
      if (a.shrinkToFit) out.stf = 1;
      if (a.textRotation === 'vertical') out.tr = { a: 0, v: 1 };
      else if (typeof a.textRotation === 'number' && a.textRotation) out.tr = { a: a.textRotation > 90 ? 90 - a.textRotation : a.textRotation };
      if (a.indent) unmapped.indents++;
    }
    if (cell.numFmt && cell.numFmt !== 'General') out.n = { pattern: cell.numFmt === 'mm-dd-yy' ? 'm/d/yy' : cell.numFmt };
    return out;
  };

  const sheetOrder = [];
  const sheets = {};
  const worksheets = workbook.worksheets;
  const nameOf = (ws) => realName.get(ws.name) ?? ws.name;
  const sheetIds = worksheets.map((_, i) => `sheet${i}`);
  const sidByName = new Map(worksheets.map((ws, i) => [nameOf(ws).toLowerCase(), sheetIds[i]]));
  const extent = new Map(sheetIds.map((sid) => [sid, { r: 0, c: 0 }]));
  const reach = (sheetName, r, c) => {
    const e = sheetName != null && extent.get(sidByName.get(sheetName.toLowerCase()));
    if (e) { e.r = Math.max(e.r, r); e.c = Math.max(e.c, c); }
  };
  let formulaCells = 0, valueCells = 0;
  worksheets.forEach((ws, index) => {
    const sid = sheetIds[index];
    sheetOrder.push(sid);
    const view = ws.views?.[0] ?? {};
    const props = ws.properties ?? {};
    const arrays = [];
    ws.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.value && typeof cell.value === 'object' && cell.value.shareType === 'array' && cell.value.ref) {
        const r = rangeOf(cell.value.ref);
        if (r) { arrays.push({ ...r, master: `${row.number - 1}:${cell.col - 1}` }); reach(nameOf(ws), r.endRow, r.endColumn); }
      }
    }));
    const inArray = (r, c) => arrays.some((a) => r >= a.startRow && r <= a.endRow && c >= a.startColumn && c <= a.endColumn && a.master !== `${r}:${c}`);
    const cellData = {};
    const rowData = {};
    let maxRow = 0, maxCol = 0;
    ws.eachRow({ includeEmpty: true }, (row) => {
      const r = row.number - 1;
      if (row.height) rowData[r] = { ...(rowData[r] ?? {}), h: Math.round(row.height * 96 / 72 * 100) / 100, ia: 0, custom: { pt: row.height } };
      if (row.hidden) rowData[r] = { ...(rowData[r] ?? {}), hd: 1 };
      (row._cells ?? []).forEach((cell) => {
        if (!cell) return;
        const c = cell.col - 1;
        const out = {};
        let v = cell.value;
        if (v && typeof v === 'object' && !(v instanceof Date)) {
          if ('formula' in v || 'sharedFormula' in v) {
            const text = cell.formula;
            if (text) { out.f = `=${normalizeFormula(text)}`; formulaCells++; scanRefs(text, (sheet, rr, cc) => reach(sheet ?? nameOf(ws), rr, cc)); }
            v = keepFormulaValues ? v.result : undefined;
          } else if (Array.isArray(v.richText)) { unmapped.richTextCells++; v = v.richText.map((run) => run.text).join(''); }
          else if ('hyperlink' in v) v = typeof v.text === 'object' ? v.text?.richText?.map((run) => run.text).join('') ?? '' : v.text;
          else if ('error' in v) v = v.error;
        }
        if (v && typeof v === 'object' && !(v instanceof Date)) v = 'error' in v ? v.error : undefined;
        if (inArray(r, c) && !out.f) v = undefined;
        if (v instanceof Date) { out.v = serial(v, date1904); out.t = 2; }
        else if (typeof v === 'number') { out.v = v; out.t = 2; }
        else if (typeof v === 'boolean') { out.v = v; out.t = 3; }
        else if (typeof v === 'string') { out.v = v; out.t = 1; }
        if (out.v !== undefined && !out.f) valueCells++;
        const style = mapStyle(cell);
        const s = styleId(style);
        if (s) out.s = s;
        if (out.v === undefined && !out.f && !out.s) return;
        (cellData[r] ??= {})[c] = out;
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      });
    });
    const columnData = {};
    const mdw = MDW[defaultFont.ff.toLowerCase()] ?? 7;
    const px = (chars) => Math.trunc(((256 * chars + Math.trunc(128 / mdw)) / 256) * mdw);
    (ws.columns ?? []).forEach((col, i) => {
      if (!col) return;
      const entry = {};
      if (col.width) { entry.w = px(col.width); entry.custom = { chars: col.width }; }
      if (col.hidden) entry.hd = 1;
      if (col.style && Object.keys(col.style).length) unmapped.columnStyles++;
      if (Object.keys(entry).length) columnData[i] = entry;
    });
    const mergeData = (ws.model?.merges ?? []).map(rangeOf).filter(Boolean);
    for (const m of mergeData) { maxRow = Math.max(maxRow, m.endRow); maxCol = Math.max(maxCol, m.endColumn); }
    for (const t of tables.get(nameOf(ws)) ?? []) reach(nameOf(ws), t.range.endRow, t.range.endColumn);
    reach(nameOf(ws), maxRow, maxCol);
    const defaultPt = props.defaultRowHeight ?? 15;
    for (let r = 0; r <= maxRow; r++) if (rowData[r]?.h === undefined) rowData[r] = { ...(rowData[r] ?? {}), h: Math.round(defaultPt * 96 / 72 * 100) / 100, ia: 0, custom: { pt: defaultPt, dflt: 1 } };
    const frozen = view.state === 'frozen' && (view.xSplit || view.ySplit);
    const defaultChars = props.defaultColWidth ?? null;
    const baseChars = Math.trunc(((props.baseColWidth ?? 8) * mdw + 5) / mdw * 256) / 256;
    sheets[sid] = {
      id: sid, name: nameOf(ws), hidden: ws.state === 'visible' ? 0 : 1, tabColor: color(props.tabColor) ?? '',
      rowCount: Math.min(MAX_ROWS, Math.max(1000, maxRow + 100)), columnCount: Math.min(MAX_COLS, Math.max(26, maxCol + 20)), zoomRatio: 1,
      defaultColumnWidth: px(defaultChars ?? baseChars), defaultRowHeight: Math.round((props.defaultRowHeight ?? 15) * 96 / 72),
      freeze: frozen ? { xSplit: view.xSplit ?? 0, ySplit: view.ySplit ?? 0, startRow: view.ySplit ?? 0, startColumn: view.xSplit ?? 0 }
        : { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 },
      showGridlines: view.showGridLines === false ? 0 : 1, rightToLeft: view.rightToLeft ? 1 : 0,
      mergeData, cellData, rowData, columnData,
      rowHeader: { width: 46 }, columnHeader: { height: 20 },
      custom: { defaultColWidthChars: defaultChars, baseColWidth: props.baseColWidth ?? 8, defaultRowHeightPt: props.defaultRowHeight ?? 15 },
    };
  });
  const scopeOf = (index) => sidByName.get(sheetsInBook[index]?.name?.toLowerCase());
  const names = definedNames(workbookXml, scopeOf);
  for (const dn of Object.values(names)) {
    const own = dn.localSheetId === 'AllDefaultWorkbook' ? null : sheets[dn.localSheetId]?.name;
    scanRefs(dn.formulaOrRefString, (sheet, r, c) => reach(sheet ?? own, r, c));
  }
  for (const sid of sheetIds) {
    const e = extent.get(sid);
    sheets[sid].rowCount = Math.min(MAX_ROWS, Math.max(1000, sheets[sid].rowCount, e.r + 1));
    sheets[sid].columnCount = Math.min(MAX_COLS, Math.max(26, sheets[sid].columnCount, e.c + 1));
  }
  const tableData = {};
  for (const [sheetName, list] of tables) { const sid = sidByName.get(sheetName.toLowerCase()); if (sid) tableData[sid] = { tables: list }; }
  const mappedTables = Object.values(tableData).reduce((n, t) => n + t.tables.length, 0);
  unmapped.tables = Math.max(0, facts.tables - mappedTables);
  const resources = [];
  if (Object.keys(names).length) resources.push({ name: 'SHEET_DEFINED_NAME_PLUGIN', data: JSON.stringify(names) });
  if (mappedTables) resources.push({ name: 'SHEET_TABLE_PLUGIN', data: JSON.stringify(tableData) });
  const data = {
    id, name: id, appVersion: '1.0.3', locale: 'enUS', styles, sheetOrder, sheets,
    defaultStyle: defaultFont, resources,
  };
  const report = { id, sheets: worksheets.map(nameOf), formulaCells, valueCells, styles: styleIds.size,
    definedNames: Object.keys(names).length, tables: mappedTables, date1904, defaultFont, unmapped,
    notLoadable: Object.entries(unmapped).filter(([k, n]) => n > 0 && ['conditionalFormatting', 'dataValidation', 'charts', 'images', 'shapes', 'comments', 'sparklines', 'pivotTables'].includes(k)).map(([k]) => k) };
  return { data, report };
}

export const isErrorText = (v) => typeof v === 'string' && ERRORS.test(v);

if (process.argv[1] && resolve(process.argv[1]) === resolve(import.meta.filename)) {
  const [input, output, reportPath] = process.argv.slice(2);
  if (!input || !output) throw new Error('usage: node xlsx2univer.mjs in.xlsx out.json [report.json]');
  const { data, report } = await convert(readFileSync(input), { id: input.split('/').at(-2) ?? 'wb' });
  writeFileSync(output, JSON.stringify(data));
  if (reportPath) writeFileSync(reportPath, JSON.stringify(report, null, 1));
  console.log(JSON.stringify({ formulaCells: report.formulaCells, notLoadable: report.notLoadable }));
}
