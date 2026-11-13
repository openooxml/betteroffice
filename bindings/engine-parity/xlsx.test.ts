import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import initWasm, * as wasmXlsx from '../../packages/xlsx/src/wasm/generated/xlsx_wasm.js';
import initPreflight, { font_preflight as wasmFontPreflight } from '../../packages/xlsx/src/wasm/generated/preflight/xlsx_preflight.js';
import { FORMATS, parseSurface } from '../../scripts/check-engine-surface';
import {
  coverageProblems,
  defineRow,
  generatedTokens,
  oneEngineExports,
  oneEngineProblems,
  printParityCounts,
  rowName,
  runKnownDifference,
  runRow,
  surfaceMethods,
  ZipReturn,
  type EngineHandle,
  type Exclusion,
  type KnownDifference,
  type OneEngineOnly,
  type Row,
} from './harness';

type XlsxEngine = typeof wasmXlsx & { font_preflight: typeof wasmFontPreflight };
type Doc = wasmXlsx.XlsxDocument;

interface Xlsx {
  engine: XlsxEngine;
  doc: Doc;
}

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006';
const MAIN_PART = 'xl/workbook.xml';
const CONTEXT = JSON.stringify({ nowSerial: 45000.75, randSeed: 42 });
const CLIENT = 4321;
const VIEWPORT = JSON.stringify({ x: 0, y: 0, width: 800, height: 400 });
const METRICS = { dpi: 96, maxDigitWidth: 7, defaultRowHeightPt: 15, fontSizePt: 11, fontFamily: 'Calibri', fontAscent: 9, fontDescent: 3 };
const at = (a1: string, sheetId = 'sheet:0') => ({ sheetId, range: { kind: 'a1', a1 } });

const rels = (entries: [string, string, string][]) =>
  `<Relationships xmlns="${PKG_REL}">${entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const text = (ref: string, value: string, style = '') => `<c r="${ref}" t="inlineStr"${style}><is><t>${value}</t></is></c>`;

const DATA_SHEET = `<worksheet xmlns="${MAIN}" xmlns:r="${REL}">` +
  '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
  '<sheetData>' +
  `<row r="1">${text('A1', 'Item', ' s="1"')}${text('B1', 'Qty', ' s="1"')}${text('C1', 'Total', ' s="1"')}</row>` +
  `<row r="2">${text('A2', 'Apples')}<c r="B2"><v>3</v></c><c r="C2"><f>B2*2</f><v>6</v></c></row>` +
  `<row r="3">${text('A3', 'Pears')}<c r="B3"><v>4</v></c><c r="C3"><f>SUM(Quantities)</f><v>7</v></c></row>` +
  `<row r="5">${text('A5', 'Merged heading')}</row>` +
  '</sheetData><mergeCells count="1"><mergeCell ref="A5:C5"/></mergeCells><drawing r:id="rIdDrawing"/></worksheet>';

const CHART = `<c:chartSpace xmlns:c="${DRAWING}/chart" xmlns:a="${DRAWING}/main" xmlns:r="${REL}"><c:chart>` +
  '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Quantities</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
  '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>' +
  '<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f></c:numRef></c:val></c:ser>' +
  '<c:axId val="1"/><c:axId val="2"/></c:barChart>' +
  '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
  '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>' +
  '</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

const ANCHOR = `<xdr:wsDr xmlns:xdr="${DRAWING}/spreadsheetDrawing" xmlns:a="${DRAWING}/main" xmlns:c="${DRAWING}/chart" xmlns:r="${REL}"><xdr:twoCellAnchor>` +
  '<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
  '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
  '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
  `<a:graphic><a:graphicData uri="${DRAWING}/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;

const STYLES = `<styleSheet xmlns="${MAIN}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>` +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>';

async function buildFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2024-01-01T00:00:00Z');
  const override = (part: string, type: string) => `<Override PartName="/${part}" ContentType="${OFFICE}.${type}"/>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      override(MAIN_PART, 'spreadsheetml.sheet.main+xml') + override('xl/worksheets/sheet1.xml', 'spreadsheetml.worksheet+xml') +
      override('xl/worksheets/sheet2.xml', 'spreadsheetml.worksheet+xml') + override('xl/styles.xml', 'spreadsheetml.styles+xml') +
      override('xl/comments1.xml', 'spreadsheetml.comments+xml') + override('xl/drawings/drawing1.xml', 'drawing+xml') +
      override('xl/charts/chart1.xml', 'drawingml.chart+xml') + '</Types>',
    '_rels/.rels': rels([['rId1', 'officeDocument', MAIN_PART]]),
    [MAIN_PART]: `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Notes" sheetId="2" r:id="rId2"/></sheets>` +
      '<definedNames><definedName name="Quantities">Data!$B$2:$B$3</definedName></definedNames></workbook>',
    'xl/_rels/workbook.xml.rels': rels([['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'worksheet', 'worksheets/sheet2.xml'], ['rId3', 'styles', 'styles.xml']]),
    'xl/styles.xml': STYLES,
    'xl/worksheets/sheet1.xml': DATA_SHEET,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rIdDrawing', 'drawing', '../drawings/drawing1.xml'], ['rIdComments', 'comments', '../comments1.xml']]),
    'xl/worksheets/sheet2.xml': `<worksheet xmlns="${MAIN}"><sheetData><row r="1">${text('A1', 'Second sheet')}</row></sheetData></worksheet>`,
    'xl/comments1.xml': `<comments xmlns="${MAIN}"><authors><author>Parity</author></authors><commentList><comment ref="B2" authorId="0"><text><r><t>Seeded comment</t></r></text></comment></commentList></comments>`,
    'xl/drawings/drawing1.xml': ANCHOR,
    'xl/drawings/_rels/drawing1.xml.rels': rels([['rIdChart', 'chart', '../charts/chart1.xml']]),
    'xl/charts/chart1.xml': CHART,
  };
  for (const [name, content] of Object.entries(parts)) zip.file(name, content, { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const wasmEngine: XlsxEngine = { ...wasmXlsx, font_preflight: wasmFontPreflight };
const nativeEngine = createRequire(import.meta.url)('../xlsx-native/index.js') as XlsxEngine;
let fixture: Uint8Array;

beforeAll(async () => {
  await initWasm({ module_or_path: readFileSync(new URL('../../packages/xlsx/src/wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)) });
  await initPreflight({ module_or_path: readFileSync(new URL('../../packages/xlsx/src/wasm/generated/preflight/xlsx_preflight_bg.wasm', import.meta.url)) });
  fixture = await buildFixture();
});

// documentVersion() is "{nonce:32 hex}-{epoch}"; the nonce is minted per open.
function versionNonce(doc: Doc): string {
  const version = doc.documentVersion();
  if (!/^[0-9a-f]{32}-\d+$/.test(version)) throw new Error(`documentVersion ${JSON.stringify(version)} is not {nonce:32 hex}-{epoch}`);
  return version.slice(0, version.lastIndexOf('-'));
}

function engineFor(engine: XlsxEngine) {
  return async (): Promise<EngineHandle<Xlsx>> => {
    const doc = engine.XlsxDocument.openWithCalculationJson(fixture, CONTEXT);
    return {
      target: { engine, doc },
      generated: () => [versionNonce(doc)],
      save: async () => doc.saveBytes(),
      free: () => doc.free(),
    };
  };
}

const makeWasm = engineFor(wasmEngine);
const makeNative = engineFor(nativeEngine);

function scoped<R>(doc: Doc, read: (doc: Doc) => R): R {
  try {
    return read(doc);
  } finally {
    doc.free();
  }
}

const json = (value: unknown) => JSON.stringify(value);
const editCell = (doc: Doc, row: number, col: number, input: string) => doc.editCellJson(json({ sheet: 0, row, col, input }));
const propose = (doc: Doc) => (JSON.parse(doc.proposeJson(json({ agentId: 'parity-agent', note: 'parity', edits: [{ sheet: 0, row: 3, col: 1, input: '5' }] }))) as { id: string }).id;
const editRequest = (doc: Doc) => json({
  expectVersion: doc.documentVersion(),
  steps: [{ op: 'setFormulas', target: at('D2'), formulas: [['C2+B3']] }, { op: 'setCellInputs', target: at('A4'), inputs: [['Plums']] }],
});
const withoutNonce = (doc: Doc, values: unknown[]) => values.map((value) => generatedTokens(value, [versionNonce(doc)]));
const collab = (engine: XlsxEngine, client = CLIENT) => engine.XlsxDocument.openCollaborative(fixture, client);
const editedCollab = (engine: XlsxEngine) => {
  const doc = collab(engine);
  editCell(doc, 1, 1, '12');
  return doc;
};
const drainAll = (doc: Doc) => {
  const events: Uint8Array[] = [];
  for (let event = doc.drainUpdateEvent(); event.byteLength > 0; event = doc.drainUpdateEvent()) events.push(event);
  return events;
};

const rows: Row<Xlsx>[] = [
  { method: 'XlsxDocument.open', call: ({ engine }) => scoped(engine.XlsxDocument.open(fixture), (d) => [d.sheetInfoJson(), d.saveWarningsJson()]) },
  {
    method: 'XlsxDocument.openWithCalculationJson',
    call: ({ engine }) => scoped(engine.XlsxDocument.openWithCalculationJson(fixture, CONTEXT), (d) => [d.sheetInfoJson(), d.calculationStatusJson()]),
  },
  { method: 'XlsxDocument.openCollaborative', call: ({ engine }) => scoped(engine.XlsxDocument.openCollaborative(fixture, CLIENT), (d) => d.sheetInfoJson()) },
  {
    method: 'XlsxDocument.openCollaborative',
    name: 'XlsxDocument.openCollaborative (client 0)',
    throws: { name: 'string', reason: 'a client id must be positive' },
    call: ({ engine }) => engine.XlsxDocument.openCollaborative(fixture, 0),
  },
  defineRow({
    method: 'XlsxDocument.openRecoveryCheckpoint',
    setup: ({ doc }: Xlsx) => {
      editCell(doc, 6, 0, 'checkpointed');
      return doc.recoveryCheckpoint(0);
    },
    call: ({ engine }, checkpoint) => scoped(engine.XlsxDocument.openRecoveryCheckpoint(checkpoint, 0), (d) => [d.sheetInfoJson(), d.cellJson(json({ sheet: 0, row: 6, col: 0 }))]),
  }),
  {
    method: 'XlsxDocument.openRecoveryCheckpoint',
    name: 'XlsxDocument.openRecoveryCheckpoint (empty)',
    throws: { name: 'Error', reason: 'empty bytes are not a checkpoint' },
    call: ({ engine }) => engine.XlsxDocument.openRecoveryCheckpoint(new Uint8Array(), 0),
  },
  { method: 'XlsxDocument.clientId', call: ({ engine }) => scoped(engine.XlsxDocument.openCollaborative(fixture, CLIENT), (d) => d.clientId) },
  { method: 'XlsxDocument.encodeStateVector', call: ({ engine }) => scoped(editedCollab(engine), (d) => d.encodeStateVector()) },
  { method: 'XlsxDocument.encodeStateAsUpdate', call: ({ engine }) => scoped(editedCollab(engine), (d) => d.encodeStateAsUpdate()) },
  {
    method: 'XlsxDocument.encodeDiff',
    call: ({ engine }) => scoped(editedCollab(engine), (d) => d.encodeDiff(scoped(collab(engine, CLIENT + 1), (peer) => peer.encodeStateVector()))),
  },
  {
    method: 'XlsxDocument.applyUpdateJson',
    call: ({ engine }) => scoped(editedCollab(engine), (d) => {
      const update = scoped(collab(engine, CLIENT + 1), (peer) => {
        editCell(peer, 2, 1, '8');
        return peer.encodeDiff(d.encodeStateVector());
      });
      return withoutNonce(d, [d.applyUpdateJson(update), d.sheetInfoJson(), d.readCellsJson(json({ ranges: [at('A1:C5')] }))]);
    }),
  },
  {
    method: 'XlsxDocument.applyUpdateJson',
    name: 'XlsxDocument.applyUpdateJson refuses a standalone document',
    throws: { name: 'string', reason: 'a standalone workbook takes no collaboration updates' },
    call: ({ engine, doc }) => doc.applyUpdateJson(scoped(editedCollab(engine), (d) => d.encodeStateAsUpdate())),
  },
  {
    method: 'XlsxDocument.startUpdateObservation',
    call: ({ engine }) => scoped(collab(engine), (d) => {
      d.startUpdateObservation();
      editCell(d, 1, 1, '12');
      return drainAll(d);
    }),
  },
  {
    method: 'XlsxDocument.drainUpdateEvent',
    call: ({ engine }) => scoped(collab(engine), (d) => {
      d.startUpdateObservation();
      editCell(d, 1, 1, '12');
      return [d.drainUpdateEvent(), d.drainUpdateEvent()];
    }),
  },
  {
    method: 'XlsxDocument.clearUpdateObservation',
    call: ({ engine }) => scoped(collab(engine), (d) => {
      d.startUpdateObservation();
      editCell(d, 1, 1, '12');
      d.clearUpdateObservation();
      return d.drainUpdateEvent();
    }),
  },
  { method: 'XlsxDocument.version', call: ({ engine }) => engine.XlsxDocument.version() },
  { method: 'XlsxDocument.saveBytes', call: ({ doc }) => new ZipReturn(doc.saveBytes()) },
  { method: 'XlsxDocument.saveWarningsJson', call: ({ doc }) => doc.saveWarningsJson() },
  { method: 'XlsxDocument.sheetInfoJson', call: ({ doc }) => doc.sheetInfoJson() },
  { method: 'XlsxDocument.sheetCount', call: ({ doc }) => doc.sheetCount() },
  { method: 'XlsxDocument.sheetInfoForJson', call: ({ doc }) => doc.sheetInfoForJson(1) },
  { method: 'XlsxDocument.cellJson', call: ({ doc }) => [0, 1, 2].map((col) => doc.cellJson(json({ sheet: 0, row: 2, col }))) },
  { method: 'XlsxDocument.readCellsJson', call: ({ doc }) => doc.readCellsJson(json({ ranges: [at('A1:C5'), at('A1', 'sheet:1')] })) },
  { method: 'XlsxDocument.listDefinedNamesJson', call: ({ doc }) => doc.listDefinedNamesJson('{}') },
  { method: 'XlsxDocument.documentVersion', call: ({ doc }) => doc.documentVersion() },
  { method: 'XlsxDocument.editCellJson', mutates: true, call: ({ doc }) => editCell(doc, 1, 1, '12') },
  { method: 'XlsxDocument.editCellJson', name: 'XlsxDocument.editCellJson (formula)', mutates: true, call: ({ doc }) => editCell(doc, 3, 2, '=SUM(B2:B3)*2') },
  {
    method: 'XlsxDocument.editCellsJson',
    mutates: true,
    call: ({ doc }) => doc.editCellsJson(json({ sheet: 0, edits: [{ row: 3, col: 0, input: 'Plums' }, { row: 3, col: 1, input: '9' }] })),
  },
  {
    method: 'XlsxDocument.patchRangeStyleJson',
    mutates: true,
    call: ({ doc }) => doc.patchRangeStyleJson(json({ sheet: 0, range: 'A2:C3', patch: { italic: true, fillColor: '#abcdef', horizontalAlignment: 'center' } })),
  },
  {
    method: 'XlsxDocument.setRangeNumberFormatJson',
    mutates: true,
    call: ({ doc }) => doc.setRangeNumberFormatJson(json({ sheet: 0, range: 'B2:B3', format: { type: 'currency' } })),
  },
  {
    method: 'XlsxDocument.applyOpsJson',
    mutates: true,
    call: ({ doc }) => doc.applyOpsJson(json({ ops: [{ type: 'mergeCells', sheet: 0, range: { start: { row: 6, col: 0 }, end: { row: 6, col: 2 } } }] })),
  },
  {
    method: 'XlsxDocument.applyOpsJson',
    name: 'XlsxDocument.applyOpsJson (insert rows)',
    mutates: true,
    call: ({ doc }) => doc.applyOpsJson(json({ ops: [{ type: 'insertRows', sheet: 0, at: 1, count: 2 }] })),
  },
  { method: 'XlsxDocument.fillRangeJson', mutates: true, call: ({ doc }) => doc.fillRangeJson(json({ sheet: 1, source: 'A1', destination: 'A1:A3' })) },
  {
    method: 'XlsxDocument.fillRangeJson',
    name: 'XlsxDocument.fillRangeJson refuses on a commented sheet',
    mutates: true,
    throws: { name: 'string', reason: 'fill cannot safely preserve comments' },
    call: ({ doc }) => doc.fillRangeJson(json({ sheet: 0, source: 'C2', destination: 'C2:C4' })),
  },
  { method: 'XlsxDocument.validateEditsJson', call: ({ doc }) => doc.validateEditsJson(editRequest(doc)) },
  { method: 'XlsxDocument.applyEditsJson', mutates: true, call: ({ doc }) => doc.applyEditsJson(editRequest(doc)) },
  defineRow({
    method: 'XlsxDocument.undoJson',
    mutates: true,
    setup: ({ doc }: Xlsx) => editCell(doc, 1, 1, '12'),
    call: ({ doc }) => doc.undoJson(),
  }),
  defineRow({
    method: 'XlsxDocument.redoJson',
    mutates: true,
    setup: ({ doc }: Xlsx) => {
      editCell(doc, 1, 1, '12');
      doc.undoJson();
    },
    call: ({ doc }) => doc.redoJson(),
  }),
  { method: 'XlsxDocument.setActiveSheet', mutates: true, call: ({ doc }) => doc.setActiveSheet(1) },
  {
    method: 'XlsxDocument.setCalculationContextJson',
    mutates: true,
    call: ({ doc }) => {
      const result = doc.setCalculationContextJson(json({ nowSerial: 46000.5, randSeed: 7 }));
      return { result, after: editCell(doc, 7, 0, '=NOW()') };
    },
  },
  { method: 'XlsxDocument.recalculateJson', mutates: true, call: ({ doc }) => doc.recalculateJson() },
  {
    method: 'XlsxDocument.translateFormulasJson',
    call: ({ doc }) => doc.translateFormulasJson(json([{ formula: 'B2*2', from: { row: 1, col: 2 }, to: { row: 4, col: 3 } }, { formula: 'A1', from: { row: 0, col: 0 }, to: { row: 0, col: 0 } }])),
  },
  { method: 'XlsxDocument.proposeJson', mutates: true, call: ({ doc }) => doc.proposeJson(json({ agentId: 'parity-agent', note: null, edits: [{ sheet: 0, row: 3, col: 1, input: '5' }] })) },
  defineRow({
    method: 'XlsxDocument.listProposalsJson',
    setup: ({ doc }: Xlsx) => propose(doc),
    call: ({ doc }) => doc.listProposalsJson(),
  }),
  defineRow({
    method: 'XlsxDocument.acceptProposalJson',
    mutates: true,
    setup: ({ doc }: Xlsx) => propose(doc),
    call: ({ doc }, id) => doc.acceptProposalJson(json({ id, force: false })),
  }),
  defineRow({
    method: 'XlsxDocument.rejectProposalJson',
    mutates: true,
    setup: ({ doc }: Xlsx) => propose(doc),
    call: ({ doc }, id) => doc.rejectProposalJson(json({ id })),
  }),
  { method: 'XlsxDocument.displayListJson', call: ({ doc }) => doc.displayListJson(VIEWPORT) },
  { method: 'XlsxDocument.displayListForJson', call: ({ doc }) => doc.displayListForJson(1, VIEWPORT) },
  { method: 'XlsxDocument.pngDisplayListJson', call: ({ doc }) => doc.pngDisplayListJson(VIEWPORT) },
  { method: 'XlsxDocument.rangePngDisplayListJson', call: ({ doc }) => [doc.rangePngDisplayListJson('{}'), doc.rangePngDisplayListJson(json({ range: 'A1:C5', scale: 2 }))] },
  {
    method: 'XlsxDocument.printDisplayListJson',
    call: ({ doc }) => doc.printDisplayListJson(json({ sheet: 0, range: 'A1:C5', metrics: METRICS, gridlines: true })),
  },
  { method: 'XlsxDocument.visibleMergedRangesJson', call: ({ doc }) => doc.visibleMergedRangesJson(0, VIEWPORT) },
  { method: 'XlsxDocument.mergedRangesJson', call: ({ doc }) => doc.mergedRangesJson(json({ sheet: 0, range: 'A1:D10' })) },
  {
    method: 'XlsxDocument.chartAtPointJson',
    call: ({ doc }) => [[350, 80], [10, 10]].map(([x, y]) => doc.chartAtPointJson(json({ viewport: JSON.parse(VIEWPORT), x, y }))),
  },
  { method: 'XlsxDocument.cellRectJson', call: ({ doc }) => doc.cellRectJson(json({ sheet: 0, row: 4, col: 1 })) },
  { method: 'XlsxDocument.cellPositionJson', call: ({ doc }) => doc.cellPositionJson(json({ sheet: 0, row: 4, col: 0 })) },
  {
    method: 'XlsxDocument.selectionDataJson',
    call: ({ doc }) => ['usedRange', 'merge', 'down'].map((direction) => doc.selectionDataJson(json({ sheet: 0, row: 4, col: 0, direction }))),
  },
  { method: 'XlsxDocument.exportStructuredJson', call: ({ doc }) => doc.exportStructuredJson('{}') },
  { method: 'XlsxDocument.exportMarkdownJson', call: ({ doc }) => doc.exportMarkdownJson('{}', '{}') },
  { method: 'exportXlsxStructuredJson', call: ({ engine }) => engine.exportXlsxStructuredJson(fixture, '{}') },
  { method: 'exportXlsxMarkdownJson', call: ({ engine }) => engine.exportXlsxMarkdownJson(fixture, '{}', '{}') },
  {
    method: 'renderXlsxMarkdownJson',
    call: ({ engine }) => engine.renderXlsxMarkdownJson(wasmEngine.exportXlsxStructuredJson(fixture, '{}'), '{}'),
  },
  { method: 'displayListCellLimit', call: ({ engine }) => engine.displayListCellLimit() },
  { method: 'font_preflight', call: ({ engine }) => engine.font_preflight(fixture) },
];

const GOLDEN = 'scripts/parity/xlsx.ts drives it on both engines against parity/xlsx/expected.json';
const TIMINGS = 'profiling twin of a covered call; only its wall-clock timings differ';
const exclusions: Exclusion[] = [
  { method: 'allocationFailure', reason: 'heap accounting of the calling thread; the native and wasm allocators differ by design' },
  { method: 'install_panic_hook', reason: 'process-wide panic logging with no document result' },
  { method: 'XlsxDocument.free', reason: 'releases the object; the harness frees every document itself' },
  { method: 'XlsxSnapshotBuilder.free', reason: 'releases the object; nothing to compare' },
  { method: 'XlsxDocument.editCellProfiledJson', reason: TIMINGS },
  { method: 'XlsxDocument.applyOpsProfiledJson', reason: TIMINGS },
  { method: 'XlsxDocument.displayListProfiledJson', reason: TIMINGS },
  { method: 'XlsxDocument.applyFormatJson', reason: GOLDEN },
  { method: 'XlsxDocument.calculationStatusJson', reason: GOLDEN },
  { method: 'XlsxDocument.captureFormatJson', reason: GOLDEN },
  { method: 'XlsxDocument.chartsJson', reason: GOLDEN },
  { method: 'XlsxDocument.conditionalFormatsJson', reason: GOLDEN },
  { method: 'XlsxDocument.editChartsJson', reason: GOLDEN },
  { method: 'XlsxDocument.findTextJson', reason: GOLDEN },
  { method: 'XlsxDocument.historyStateJson', reason: GOLDEN },
  { method: 'XlsxDocument.moveChartJson', reason: GOLDEN },
  { method: 'XlsxDocument.moveRangeJson', reason: GOLDEN },
  { method: 'XlsxDocument.rangeCellsJson', reason: GOLDEN },
  { method: 'XlsxDocument.searchTextJson', reason: GOLDEN },
  { method: 'XlsxDocument.selectionFormattingJson', reason: GOLDEN },
];

const knownDifferences: KnownDifference[] = [];

const notYetCovered: string[] = [
  'XlsxDocument.adoptPeerVersion', 'XlsxDocument.beginPeerSnapshot',
  'XlsxDocument.endPeerSnapshot', 'XlsxDocument.nextPeerSnapshotChunk', 'XlsxDocument.openWithPeerHydrationJson', 'XlsxDocument.peerHydrationJson',
  'XlsxDocument.recoveryCheckpoint', 'XlsxDocument.resizeChartJson',
  'XlsxSnapshotBuilder.advance', 'XlsxSnapshotBuilder.finish', 'XlsxSnapshotBuilder.push',
  'decodeFontFace',
];

const NO_RASTERIZER = 'the wasm build carries no rasterizer; wasm draws PNG targets through the canvas in packages/xlsx/src/wasm/enginePng.ts';
const FONT_PACKS = 'native reads installed font pack directories; wasm has no filesystem and loads @betteroffice/fonts bytes in JS (packages/xlsx/src/wasm/enginePngFonts.ts:11)';
const oneEngineOnly: OneEngineOnly[] = [
  { name: 'XlsxDocument.renderPngTarget', side: 'native', reason: 'BY_DESIGN', note: NO_RASTERIZER },
  { name: 'configureFontPacks', side: 'native', reason: 'BY_DESIGN', note: FONT_PACKS },
  { name: 'fontPackDiagnosticsJson', side: 'native', reason: 'BY_DESIGN', note: FONT_PACKS },
  { name: 'engineHeapLiveBytes', side: 'native', reason: 'PARKED', note: 'heap accounting of the native allocator; no wasm counterpart yet' },
  { name: 'setHeapLimit', side: 'native', reason: 'PARKED', note: 'native heap cap; no wasm counterpart yet' },
  { name: 'engineVersion', side: 'native', reason: 'PARKED', note: 'engine build version; not exported by the wasm build yet' },
  { name: '__napiBindingTarget', side: 'native', reason: 'BUILD_METADATA', note: 'napi-rs build target marker' },
];

function xlsxSurfaces() {
  const read = (path: string) => parseSurface(readFileSync(new URL(`../../${path}`, import.meta.url), 'utf8'));
  const wasm = Object.entries(FORMATS).filter(([name]) => name === 'xlsx' || name.startsWith('xlsx-')).map(([, format]) => read(format.wasm));
  const native = read(FORMATS.xlsx!.native);
  return { surface: surfaceMethods(wasm, native), oneEngine: oneEngineExports(wasm, native) };
}

describe('xlsx native and wasm parity table', () => {
  test('every method on both engines is a row, an exclusion or not yet covered', () => {
    const { surface, oneEngine } = xlsxSurfaces();
    const plan = { rows, exclusions, notYetCovered, knownDifferences, oneEngineOnly };
    printParityCounts('xlsx', surface, plan);
    expect(coverageProblems(surface, plan)).toEqual([]);
    expect(oneEngineProblems(oneEngine, oneEngineOnly)).toEqual([]);
  });

  test('the not yet covered list only shrinks', () => {
    expect(notYetCovered.length).toBe(12);
  });

  for (const row of rows) {
    const known = knownDifferences.find((entry) => entry.method === row.method);
    test(known ? `${rowName(row)} still differs: ${known.reason}` : `${rowName(row)} matches`, async () => {
      await (known ? runKnownDifference(row, makeWasm, makeNative, MAIN_PART) : runRow(row, makeWasm, makeNative, MAIN_PART));
    }, 30_000);
  }
});
