// Formula columns are 14 characters (103 px) wide in both arms so the full chain total paints and a recalc changes pixels.
// Deterministic synthetic workbooks. Base shapes: `data` number columns (some text), then SUM of the row, then a running
// total of that sum (the chain, the last column). Dense shape: A-H data, I = SUM(A:H), J..L dependent per-row formulas,
// M = running total of L (the chain), N1:N10 = SUMIFS over the whole L/A ranges.
// `bun gen.mjs <outDir> w100k,r1m,...` writes <name>.xlsx (for bo-xlsx) and <name>.json (IWorkbookData for univer-sheets)
// and prints the sha256 of both.
import { deflateRawSync } from 'node:zlib';
import { createHash } from 'node:crypto';
import { closeSync, mkdirSync, openSync, writeFileSync, writeSync } from 'node:fs';
import { resolve } from 'node:path';

export const SHAPES = {
  w100k: { rows: 10_000, data: 8 }, w300k: { rows: 30_000, data: 8 }, w1m: { rows: 100_000, data: 8 },
  r100k: { rows: 100_000, data: 8 }, r300k: { rows: 300_000, data: 8 }, r1m: { rows: 1_000_000, data: 8 },
  wide50k: { rows: 50_000, data: 198 },
  dense100k: { rows: 100_000, data: 8, dense: true },
};
export const WORKBOOKS = Object.fromEntries(Object.entries(SHAPES).map(([name, s]) => [name, s.rows]));
export const SUMIFS_AT = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map((k) => k * 100);

export const colName = (c) => { let s = ''; for (c += 1; c; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; };

/** Column indexes of one shape: the row SUM, the running-total chain, the optional SUMIFS column, and the column count. */
export function layout(name) {
  const s = SHAPES[name];
  if (!s) throw new Error(`unknown workbook ${name}`);
  if (s.dense) return { ...s, sumCol: 8, chainCol: 12, sumifsCol: 13, cols: 14 };
  return { ...s, sumCol: s.data, chainCol: s.data + 1, sumifsCol: null, cols: s.data + 2 };
}

function hash(r, c) {
  let x = Math.imul(r + 1, 0x9e3779b1) ^ Math.imul(c + 7, 0x85ebca6b);
  x = Math.imul(x ^ (x >>> 16), 0x7feb352d);
  x = Math.imul(x ^ (x >>> 15), 0x846ca68b);
  return (x ^ (x >>> 16)) >>> 0;
}

/** Zero-based row and data column: a number 0..999, or a short text in about 2% of cells. */
export function value(r, c) {
  const h = hash(r, c);
  return h % 53 === 0 ? `t${r + 1}${colName(c)}` : h % 1000;
}

export const numeric = (v) => (typeof v === 'number' ? v : 0);

export function rowSum(r) {
  let s = 0;
  for (let c = 0; c < 8; c++) s += numeric(value(r, c));
  return s;
}

export function total(rows) {
  let s = 0;
  for (let r = 0; r < rows; r++) s += rowSum(r);
  return s;
}

/** Formula cells of row r given the data getter, with their values; `prev` is the chain value of row r-1. */
function formulas(L, r, get, prev) {
  const n = r + 1;
  const C = colName;
  let s = 0;
  for (let c = 0; c < L.data; c++) s += numeric(get(r, c));
  const out = [{ c: L.sumCol, f: `SUM(A${n}:${C(L.data - 1)}${n})`, v: s }];
  let link = s;
  if (L.dense) {
    const a = numeric(get(r, 0)), b = numeric(get(r, 1));
    const j = 2 * s - a;
    const k = j > 3 * b ? j - b : b;
    link = k + (s % 7);
    out.push({ c: 9, f: `I${n}*2-SUM(A${n})`, v: j });
    out.push({ c: 10, f: `IF(J${n}>SUM(B${n})*3,J${n}-SUM(B${n}),SUM(B${n}))`, v: k });
    out.push({ c: 11, f: `K${n}+MOD(I${n},7)`, v: link });
  }
  const chain = prev + link;
  const from = C(L.chainCol - 1);
  out.push({ c: L.chainCol, f: n === 1 ? `${from}1` : `${C(L.chainCol)}${n - 1}+${from}${n}`, v: chain });
  return { out, chain, link };
}

/** Values the workbook must show for the given data getter: the last chain cell and (dense) N1:N10. */
export function expected(name, get = value) {
  const L = layout(name);
  let chain = 0;
  const sumifs = L.dense ? SUMIFS_AT.map(() => 0) : null;
  for (let r = 0; r < L.rows; r++) {
    const row = formulas(L, r, get, chain);
    chain = row.chain;
    if (sumifs) {
      const a = get(r, 0);
      if (typeof a === 'number') for (let k = 0; k < SUMIFS_AT.length; k++) if (a >= SUMIFS_AT[k]) sumifs[k] += row.link;
    }
  }
  return { chain, sumifs };
}

function sumifsCells(L, r, sumifs) {
  if (L.sumifsCol === null || r >= SUMIFS_AT.length) return [];
  return [{ c: L.sumifsCol, f: `SUMIFS($L$1:$L$${L.rows},$A$1:$A$${L.rows},">=${SUMIFS_AT[r]}")`, v: sumifs[r] }];
}

const xmlEscape = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function sheetXml(name, sumifs) {
  const L = layout(name);
  const strings = [];
  const index = new Map();
  const chunks = [];
  let part = [`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${colName(L.cols - 1)}${L.rows}"/><cols><col min="${L.sumCol + 1}" max="${L.cols}" width="14" customWidth="1"/></cols><sheetData>`];
  let chain = 0;
  for (let r = 0; r < L.rows; r++) {
    const n = r + 1;
    let row = `<row r="${n}">`;
    for (let c = 0; c < L.data; c++) {
      const v = value(r, c);
      if (typeof v === 'number') row += `<c r="${colName(c)}${n}"><v>${v}</v></c>`;
      else {
        let i = index.get(v);
        if (i === undefined) { i = strings.length; strings.push(v); index.set(v, i); }
        row += `<c r="${colName(c)}${n}" t="s"><v>${i}</v></c>`;
      }
    }
    const f = formulas(L, r, value, chain);
    chain = f.chain;
    for (const cell of [...f.out, ...sumifsCells(L, r, sumifs)]) row += `<c r="${colName(cell.c)}${n}"><f>${xmlEscape(cell.f)}</f><v>${cell.v}</v></c>`;
    part.push(row, '</row>');
    if (part.length > 20_000) { chunks.push(Buffer.from(part.join(''), 'utf8')); part = []; }
  }
  part.push('</sheetData></worksheet>');
  chunks.push(Buffer.from(part.join(''), 'utf8'));
  return { sheet: Buffer.concat(chunks), strings };
}

const CRC = new Int32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c;
});
function crc32(bytes) {
  let c = -1;
  for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function zip(files) {
  const chunks = [];
  const central = [];
  let offset = 0;
  for (const [name, text] of files) {
    const raw = typeof text === 'string' ? Buffer.from(text, 'utf8') : text;
    const data = deflateRawSync(raw, { level: 6 });
    const nameBytes = Buffer.from(name, 'utf8');
    const crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0, 6); local.writeUInt16LE(8, 8);
    local.writeUInt16LE(0, 10); local.writeUInt16LE(0x21, 12); local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18); local.writeUInt32LE(raw.length, 22); local.writeUInt16LE(nameBytes.length, 26);
    chunks.push(local, nameBytes, data);
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50, 0); entry.writeUInt16LE(20, 4); entry.writeUInt16LE(20, 6); entry.writeUInt16LE(0, 8);
    entry.writeUInt16LE(8, 10); entry.writeUInt16LE(0, 12); entry.writeUInt16LE(0x21, 14); entry.writeUInt32LE(crc, 16);
    entry.writeUInt32LE(data.length, 20); entry.writeUInt32LE(raw.length, 24); entry.writeUInt16LE(nameBytes.length, 28);
    entry.writeUInt32LE(offset, 42);
    central.push(entry, nameBytes);
    offset += 30 + nameBytes.length + data.length;
  }
  const size = central.reduce((s, b) => s + b.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(size, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, ...central, end]);
}

export function xlsx(name, sumifs = expected(name).sumifs) {
  const { sheet, strings } = sheetXml(name, sumifs);
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  return zip([
    ['[Content_Types].xml', `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`],
    ['_rels/.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `${head}<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`],
    ['xl/_rels/workbook.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${rel}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`],
    ['xl/styles.xml', `${head}<styleSheet xmlns="${ns}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`],
    ['xl/sharedStrings.xml', `${head}<sst xmlns="${ns}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si><t>${xmlEscape(s)}</t></si>`).join('')}</sst>`],
    ['xl/worksheets/sheet1.xml', sheet],
  ]);
}

/**
 * Univer IWorkbookData with the same cells (cached values plus `f` formulas; t: 1 string, 2 number), streamed to `path`
 * because the big shapes exceed a comfortable single string. Returns the byte count and sha256.
 */
export function writeSnapshot(name, path, sumifs = expected(name).sumifs) {
  const L = layout(name);
  const fd = openSync(path, 'w');
  const sha = createHash('sha256');
  let bytes = 0, buffer = [], pending = 0;
  const put = (s) => { buffer.push(s); pending += s.length; if (pending > 1 << 22) flush(); };
  const flush = () => { const b = Buffer.from(buffer.join(''), 'utf8'); writeSync(fd, b); sha.update(b); bytes += b.length; buffer = []; pending = 0; };
  const columnData = {};
  for (let c = L.sumCol; c < L.cols; c++) columnData[c] = { w: 103 };
  put(`{"id":"cmp-${name}","name":"${name}","appVersion":"1.0.3","locale":"enUS","styles":{},"sheetOrder":["sheet1"],"sheets":{"sheet1":{"id":"sheet1","name":"Sheet1","rowCount":${L.rows + 20},"columnCount":${Math.max(20, L.cols + 10)},"columnData":${JSON.stringify(columnData)},"cellData":{`);
  let chain = 0;
  for (let r = 0; r < L.rows; r++) {
    let row = `${r ? ',' : ''}"${r}":{`;
    for (let c = 0; c < L.data; c++) {
      const v = value(r, c);
      row += `${c ? ',' : ''}"${c}":${typeof v === 'number' ? `{"v":${v},"t":2}` : `{"v":${JSON.stringify(v)},"t":1}`}`;
    }
    const f = formulas(L, r, value, chain);
    chain = f.chain;
    for (const cell of [...f.out, ...sumifsCells(L, r, sumifs)]) row += `,"${cell.c}":{"f":${JSON.stringify(`=${cell.f}`)},"v":${cell.v},"t":2}`;
    put(`${row}}`);
  }
  put('}}}}');
  flush();
  closeSync(fd);
  return { bytes, sha256: sha.digest('hex') };
}

if (import.meta.main ?? process.argv[1]?.endsWith('gen.mjs')) {
  const [outDir, list] = process.argv.slice(2);
  if (!outDir || !list) throw new Error(`usage: gen.mjs <outDir> ${Object.keys(SHAPES).join(',')}`);
  mkdirSync(outDir, { recursive: true });
  const manifest = {};
  for (const name of list.split(',')) {
    const L = layout(name);
    const want = expected(name);
    const bytes = xlsx(name, want.sumifs);
    writeFileSync(resolve(outDir, `${name}.xlsx`), bytes);
    const xlsxSha = createHash('sha256').update(bytes).digest('hex');
    console.log(`sha256 ${name}.xlsx ${xlsxSha}`);
    const json = writeSnapshot(name, resolve(outDir, `${name}.json`), want.sumifs);
    console.log(`sha256 ${name}.json ${json.sha256}`);
    manifest[name] = { rows: L.rows, cols: L.cols, cells: L.rows * L.cols, xlsxBytes: bytes.length, xlsxSha, jsonBytes: json.bytes, jsonSha: json.sha256,
      chainCell: `${colName(L.chainCol)}${L.rows}`, chain: want.chain, sumifs: want.sumifs };
    console.log(`gen ${name} rows=${L.rows} cols=${L.cols} xlsx=${bytes.length} json=${json.bytes} chain=${want.chain}`);
  }
  writeFileSync(resolve(outDir, 'manifest.json'), JSON.stringify(manifest, null, 1));
}
