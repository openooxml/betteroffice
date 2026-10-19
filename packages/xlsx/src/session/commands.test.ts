import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { SessionFailure } from '../../../../shared/office-session';
import { isTransferResult } from '../../../../shared/office-session/protocol';
import * as wasm from '../wasm/loader';
import type { XlsxEditResult, XlsxRangeTarget } from '../edits';
import type {
  WorkbookAck, WorkbookCommand, WorkbookCommandPayload, WorkbookFormatTokenResult, WorkbookToolbarPayload,
} from './commandProtocol';
import type { WorkbookSession } from './client';
import { classifyRefusal, createWorkbookCommandState } from './commandHandlers';
import { createWorkbookCommandHost } from './commandHost';
import * as toolbarCommands from './commandToolbar';
import { WORKBOOK_REPLAY_MUTATORS, type WorkbookReplayMethod } from './replay';
import { batchRequests, capturedWorkbookCommand, createTestWorkbookSession } from './testHelpers';

let bytes: Uint8Array;
let plainBytes: Uint8Array;

beforeAll(async () => {
  await wasm.initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells><drawing r:id="rId1"/></worksheet>');
  zip.file('xl/worksheets/sheet2.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  zip.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>');
  zip.file('xl/drawings/drawing1.xml', '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>');
  zip.file('xl/drawings/_rels/drawing1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>');
  zip.file('xl/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Data!$A$1:$B$1</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>');
  bytes = await zip.generateAsync({ type: 'uint8array' });
  const plain = new JSZip();
  plain.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  plain.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  plain.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets></workbook>');
  plain.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>');
  plain.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData></worksheet>');
  plain.file('xl/worksheets/sheet2.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
  plainBytes = await plain.generateAsync({ type: 'uint8array' });
});

function command<K extends WorkbookCommand['kind']>(
  session: WorkbookSession, seq: number, kind: K, payload: unknown, afterSeq?: number
): WorkbookCommand {
  return { ...session.commandIdentity, seq, basis: session.confirmedBasis, kind, payload, afterSeq } as WorkbookCommand;
}

const selected = { sheetId: 'sheet:0', range: { kind: 'a1' as const, a1: 'A1' } };

async function trackedSession(calculation?: wasm.WorkbookCalculationContext) {
  const original = wasm.openWorkbook;
  let handle: wasm.WorkbookHandle | undefined;
  const opening = spyOn(wasm, 'openWorkbook').mockImplementation((...args) => {
    handle = original(...args);
    return handle;
  });
  try {
    const session = await createTestWorkbookSession(bytes, undefined, { calculation });
    if (!handle) throw new Error('Missing worker handle');
    return { session, handle };
  } finally { opening.mockRestore(); }
}

describe('workbook commands', () => {
  test('toolbar payload types require range targets and omit refresh targets', () => {
    type RangeToolbar = Exclude<WorkbookToolbarPayload, { id: 'refreshProposals' }>;
    type TargetedRefresh = { id: 'refreshProposals'; args: null; target: XlsxRangeTarget };
    const requiresTarget: RangeToolbar extends { target: XlsxRangeTarget } ? true : false = true;
    const refreshCannotTarget: TargetedRefresh extends WorkbookToolbarPayload ? false : true = true;
    const range: WorkbookToolbarPayload = { id: 'bold', args: null, target: selected };
    const refresh: WorkbookToolbarPayload = { id: 'refreshProposals', args: null };
    expect(requiresTarget).toBe(true);
    expect(refreshCannotTarget).toBe(true);
    expect(range.target).toEqual(selected);
    expect(refresh.target).toBeUndefined();
  });

  test('format token result type requires a string', () => {
    const requiresString: WorkbookFormatTokenResult['formatToken'] extends string ? true : false = true;
    expect(requiresString).toBe(true);
  });

  test('toolbar clear matches legacy clear and is one undo step', async () => {
    const { session, handle } = await trackedSession();
    const target = { ...selected, range: { kind: 'rowCol' as const,
      start: { row: 0, col: 0 }, end: { row: 1, col: 1 } } };
    try {
      await session.submit(command(session, 1, 'patchRangeStyle', {
        args: [0, 'A1:B2', { bold: true, fillColor: '#abcdef', textWrapping: 'wrap' }],
      }));
      await session.submit(command(session, 2, 'setNumberFormat', { args: [0, 'A1:B2', 'currency'] }));
      await session.submit(command(session, 3, 'editCells', { args: [0, [
        { row: 1, col: 0, input: 'text' }, { row: 1, col: 1, input: '=A1+1' },
        { row: 1, col: 2, input: 'outside' },
      ]] }));
      const before = await session.save();
      const history = handle.historyState();
      const formatting = handle.captureFormat(0, 'A1:B2');
      const legacy = wasm.openWorkbook(bytes);
      const clearing = spyOn(handle, 'editCells');
      try {
        legacy.patchRangeStyle(0, 'A1:B2', { bold: true, fillColor: '#abcdef', textWrapping: 'wrap' });
        legacy.setNumberFormat(0, 'A1:B2', 'currency');
        legacy.editCells(0, [
          { row: 1, col: 0, input: 'text' }, { row: 1, col: 1, input: '=A1+1' },
          { row: 1, col: 2, input: 'outside' },
        ]);
        const edits = [
          { row: 0, col: 0, input: '' }, { row: 0, col: 1, input: '' },
          { row: 1, col: 0, input: '' }, { row: 1, col: 1, input: '' },
        ];
        legacy.editCells(0, edits);
        expect(await session.submit(command(session, 4, 'applyOps', {
          toolbar: { id: 'clear', args: null, target }, selection: target,
        }))).toMatchObject({ status: 'applied', committed: true });
        expect(clearing).toHaveBeenCalledTimes(1);
        expect(clearing).toHaveBeenCalledWith(0, edits);
        expect(await session.save()).toEqual(legacy.save());
        expect(handle.captureFormat(0, 'A1:B2')).toEqual(formatting);
        expect(handle.cell(0, 1, 2).input).toBe('outside');
        expect(handle.historyState().undoDepth).toBe(history.undoDepth + 1);
        expect(session.selection).toEqual(target);
        expect(session.editState.formatting).toEqual(handle.selectionFormatting(0, 'A1:B2'));
        expect(await session.submit(command(session, 5, 'undo', { args: [] })))
          .toMatchObject({ status: 'applied', committed: true });
        expect(handle.historyState().undoDepth).toBe(history.undoDepth);
        expect(await session.save()).toEqual(before);
        expect(await session.submit(command(session, 6, 'redo', { args: [] })))
          .toMatchObject({ status: 'applied', committed: true });
        expect(await session.save()).toEqual(legacy.save());
      } finally { clearing.mockRestore(); legacy.dispose(); }
    } finally { await session.dispose(); }
  });

  test('toolbar clear preserves legacy admission refusals', async () => {
    const { session, handle } = await trackedSession();
    const clearing = spyOn(handle, 'editCells');
    try {
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const selection = structuredClone(session.selection);
      let seq = 0;
      for (const [target, reason] of [
        [{ ...selected, range: { kind: 'a1' as const, a1: 'A1:E1' } }, 'locked-target'],
        [{ ...selected, sheetId: 'missing' }, 'target-sheet-missing'],
        [{ ...selected, range: { kind: 'a1' as const, a1: 'A0' } }, 'missing-target'],
      ] as const) {
        expect(await session.submit(command(session, ++seq, 'applyOps', {
          toolbar: { id: 'clear', args: null, target },
        }))).toMatchObject({ status: 'refused', reason, committed: false });
        expect(handle.version()).toBe(version);
        expect(handle.historyState()).toEqual(history);
        expect(session.selection).toEqual(selection);
        expect(await session.save()).toEqual(before);
      }
      await session.submit(command(session, ++seq, 'setReadOnly', { readOnly: true }));
      expect(await session.submit(command(session, ++seq, 'applyOps', {
        toolbar: { id: 'clear', args: null, target: selected },
      }))).toMatchObject({ status: 'refused', reason: 'read-only', committed: false });
      expect(clearing).not.toHaveBeenCalled();
      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(await session.save()).toEqual(before);
    } finally { clearing.mockRestore(); await session.dispose(); }
  });

  test('toolbar clear refuses oversized ranges before expanding cells', async () => {
    const { session, handle } = await trackedSession();
    const resolving = spyOn(toolbarCommands, 'resolveWorkbookToolbar').mockImplementation(() => {
      throw new Error('Oversized clear reached cell expansion');
    });
    const clearing = spyOn(handle, 'editCells');
    try {
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const selection = structuredClone(session.selection);
      let seq = 0;
      for (const a1 of ['A2:A100002', 'A1:XFD1048576']) {
        const target = { ...selected, range: { kind: 'a1' as const, a1 } };
        expect(await session.submit(command(session, ++seq, 'applyOps', {
          toolbar: { id: 'clear', args: null, target }, selection: target,
        }))).toMatchObject({ status: 'refused', reason: 'limit-exceeded', committed: false });
        expect(resolving).not.toHaveBeenCalled();
        expect(clearing).not.toHaveBeenCalled();
        expect(handle.version()).toBe(version);
        expect(handle.historyState()).toEqual(history);
        expect(session.selection).toEqual(selection);
      }
      expect(await session.save()).toEqual(before);
    } finally { resolving.mockRestore(); clearing.mockRestore(); await session.dispose(); }
  });

  test('refresh proposals returns proposals without mutation', async () => {
    const { session, handle } = await trackedSession();
    try {
      expect(await session.submit(command(session, 1, 'applyOps', {
        toolbar: { id: 'refreshProposals', args: null },
      }))).toMatchObject({ status: 'applied', committed: false, result: { proposals: [] } });
      const proposed = await session.submit(command(session, 2, 'propose', {
        args: ['agent', 'Review', [{ sheet: 0, row: 0, col: 0, input: '7' }]],
      }));
      expect(proposed).toMatchObject({ status: 'applied', committed: false });
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const selection = structuredClone(session.selection);
      const proposals = handle.listProposals();
      expect(proposals).toHaveLength(1);
      const status = await session.submit(command(session, 3, 'applyOps', {
        toolbar: { id: 'refreshProposals', args: null },
      }));
      expect(status).toEqual({ seq: 3, status: 'applied', committed: false, result: { proposals } });
      expect(handle.version()).toBe(version);
      expect(session.confirmedBasis.version).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(handle.listProposals()).toEqual(proposals);
      expect(session.selection).toEqual(selection);
      expect(await session.save()).toEqual(before);
      const destination = { ...selected, range: { kind: 'a1' as const, a1: 'B1' } };
      expect(await session.submit(command(session, 4, 'applyOps', {
        toolbar: { id: 'refreshProposals', args: null, target: destination }, selection: destination,
      }))).toMatchObject({ seq: 4, status: 'refused', reason: 'invalid-command', committed: false });
      expect(session.selection).toEqual(selection);
      expect(handle.version()).toBe(version);
      expect(session.confirmedBasis.version).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(handle.listProposals()).toEqual(proposals);
      expect(await session.save()).toEqual(before);
    } finally { await session.dispose(); }
  });

  test('refresh proposals returns proposals in read-only sessions without mutation', async () => {
    const { session, handle } = await trackedSession();
    try {
      expect(await session.submit(command(session, 1, 'propose', {
        args: ['agent', 'Review', [{ sheet: 0, row: 0, col: 0, input: '7' }]],
      }))).toMatchObject({ status: 'applied', committed: false });
      expect(await session.submit(command(session, 2, 'setReadOnly', { readOnly: true })))
        .toMatchObject({ status: 'applied', committed: false });
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const selection = structuredClone(session.selection);
      const editState = structuredClone(session.editState);
      const proposals = handle.listProposals();
      const destination = { ...selected, range: { kind: 'a1' as const, a1: 'B1' } };
      expect(proposals).toHaveLength(1);
      expect(editState.readOnly).toBe(true);
      expect(await session.submit(command(session, 3, 'applyOps', {
        toolbar: { id: 'refreshProposals', args: null }, selection: destination,
      }))).toEqual({ seq: 3, status: 'applied', committed: false, result: { proposals } });
      expect(await session.submit(command(session, 4, 'applyOps', {
        toolbar: { id: 'bold', args: null, target: destination }, selection: destination,
      }))).toMatchObject({ status: 'refused', reason: 'read-only', committed: false });
      expect(handle.version()).toBe(version);
      expect(session.confirmedBasis.version).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(handle.listProposals()).toEqual(proposals);
      expect(session.selection).toEqual(selection);
      expect(session.editState).toEqual(editState);
      expect(await session.save()).toEqual(before);
    } finally { await session.dispose(); }
  });

  test('refresh proposals refuses a target-less form for any other toolbar id', async () => {
    const { session, handle } = await trackedSession();
    try {
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const selection = structuredClone(session.selection);
      let seq = 0;
      for (const [id, args] of [
        ['clear', null], ['bold', null], ['italic', null], ['strikethrough', null], ['paintFormat', null],
        ['fontFamily', { family: 'Arial' }], ['fontSize', { points: 12 }],
        ['fontSizeStep', { direction: 'increase' }], ['decimalPlaces', { direction: 'decrease' }],
        ['textColor', { color: '#123456' }], ['fillColor', { color: '#abcdef' }],
        ['numberFormat', { value: 'number' }], ['borderPreset', { value: 'all' }],
        ['borderStyle', { value: 'solid' }], ['borderColor', { color: '#000000' }],
        ['merge', null], ['merge', { value: 'all' }], ['horizontalAlignment', { value: 'center' }],
        ['verticalAlignment', { value: 'middle' }], ['textWrapping', null], ['textWrapping', { value: 'wrap' }],
        ['searchMenus', null], ['save', null], ['exportPng', null], ['print', null],
        ['undo', null], ['redo', null], ['zoom', { scale: 1 }], ['proposalsPanel', null],
        ['proposalAccept', { proposalId: 'p1' }], ['proposalReject', { proposalId: 'p1' }],
        ['underline', null], ['indent', null], ['unknown', null],
        ['refreshProposals', {}], ['refreshProposals', undefined],
      ] as const) {
        expect(await session.submit(command(session, ++seq, 'applyOps', { toolbar: { id, args } })))
          .toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
        expect(handle.version()).toBe(version);
        expect(handle.historyState()).toEqual(history);
        expect(session.selection).toEqual(selection);
      }
      expect(await session.save()).toEqual(before);
      expect(await session.submit(command(session, ++seq, 'applyOps', {
        toolbar: { id: 'refreshProposals', args: null },
      }))).toMatchObject({ status: 'applied', committed: false, result: { proposals: [] } });
    } finally { await session.dispose(); }
  });

  test('rapid relative toggles resolve in worker', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      await session.submit(command(session, 1, 'patchRangeStyle', {
        args: [0, 'A1', { bold: false, fontSize: 10, textWrapping: 'overflow' }],
      }));
      const toggle = {
        toolbar: { id: 'bold', args: null, target: selected }, selection: selected,
      };
      const first = command(session, 2, 'applyOps', toggle);
      const second = command(session, 3, 'applyOps', toggle, 2);
      const applyingFirst = session.submit(first);
      const applyingSecond = session.submit(second);
      expect(await applyingFirst).toMatchObject({ status: 'applied', committed: true });
      expect(await applyingSecond).toMatchObject({ status: 'applied', committed: true });
      expect(session.editState.formatting?.bold).toBe(false);
      const step = {
        toolbar: { id: 'fontSizeStep', args: { direction: 'increase' }, target: selected }, selection: selected,
      };
      const increase = command(session, 4, 'applyOps', step);
      const increaseAgain = command(session, 5, 'applyOps', step, 4);
      const increasing = session.submit(increase);
      const increasingAgain = session.submit(increaseAgain);
      expect(await increasing).toMatchObject({ status: 'applied', committed: true });
      expect(await increasingAgain).toMatchObject({ status: 'applied', committed: true });
      const wrap = { toolbar: { id: 'textWrapping', args: null, target: selected }, selection: selected };
      const wrapping = command(session, 6, 'applyOps', wrap);
      const wrappingAgain = command(session, 7, 'applyOps', wrap, 6);
      const applyingWrap = session.submit(wrapping);
      const applyingWrapAgain = session.submit(wrappingAgain);
      expect(await applyingWrap).toMatchObject({ status: 'applied', committed: true });
      expect(await applyingWrapAgain).toMatchObject({ status: 'applied', committed: true });
      const absolute: WorkbookCommandPayload<'patchRangeStyle'> = { args: [0, 'B1', { bold: true }] };
      const set = command(session, 8, 'patchRangeStyle', absolute);
      const setAgain = command(session, 9, 'patchRangeStyle', absolute, 8);
      const setting = session.submit(set);
      const settingAgain = session.submit(setAgain);
      expect((await setting).status).toBe('applied');
      expect(['applied', 'noop']).toContain((await settingAgain).status);
      const saved = wasm.openWorkbook(await session.save());
      try {
        expect(saved.selectionFormatting(0, 'A1')).toMatchObject({ bold: false, fontSize: 12, textWrapping: 'overflow' });
        expect(saved.selectionFormatting(0, 'B1').bold).toBe(true);
      } finally { saved.dispose(); }
    } finally { await session.dispose(); }
  });

  test('toolbar action refuses unknown id without mutation', async () => {
    const { session, handle } = await trackedSession();
    try {
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      expect(await session.submit(command(session, 1, 'applyOps', {
        toolbar: { id: 'unknown', args: null, target: selected }, selection: selected,
      }))).toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(await session.save()).toEqual(before);
      expect(await session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '9'] })))
        .toMatchObject({ status: 'applied' });
    } finally { await session.dispose(); }
  });

  test('format token round trip applies captured format', async () => {
    const { session, handle } = await trackedSession();
    const source = { ...selected, range: { kind: 'a1' as const, a1: 'B1' } };
    try {
      await session.submit(command(session, 1, 'patchRangeStyle', {
        args: [0, 'B1', { bold: true, italic: true, fontSize: 18 }],
      }));
      const version = handle.version();
      const history = handle.historyState();
      const captured = await session.submit(command(session, 2, 'applyOps', {
        toolbar: { id: 'paintFormat', args: null, target: source }, selection: source,
      }));
      expect(captured).toMatchObject({ status: 'applied', committed: false });
      const { formatToken } = captured.result as WorkbookFormatTokenResult;
      expect(typeof formatToken).toBe('string');
      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      await session.submit(command(session, 3, 'patchRangeStyle', { args: [0, 'B1', { bold: false }] }));
      expect(await session.submit(command(session, 4, 'applyFormat', {
        target: selected, formatToken, selection: selected,
      }))).toMatchObject({ status: 'applied', committed: true });
      expect(session.editState.formatting).toMatchObject({ bold: true, italic: true, fontSize: 18 });
      const saved = wasm.openWorkbook(await session.save());
      try { expect(saved.selectionFormatting(0, 'A1')).toMatchObject({ bold: true, italic: true, fontSize: 18 }); }
      finally { saved.dispose(); }
      const recaptured = await session.submit(command(session, 5, 'applyOps', {
        toolbar: { id: 'paintFormat', args: null, target: selected },
      }));
      expect(recaptured).toMatchObject({ status: 'applied', committed: false });
      expect(typeof (recaptured.result as WorkbookFormatTokenResult).formatToken).toBe('string');
      expect((recaptured.result as WorkbookFormatTokenResult).formatToken).not.toBe(formatToken);
    } finally { await session.dispose(); }
  });

  test('peerless format capture works without crypto.randomUUID', async () => {
    const descriptor = Object.getOwnPropertyDescriptor(globalThis.crypto, 'randomUUID');
    Object.defineProperty(globalThis.crypto, 'randomUUID', { configurable: true, value: undefined });
    let session: WorkbookSession | undefined;
    try {
      expect(typeof globalThis.crypto.getRandomValues).toBe('function');
      const tracked = await trackedSession();
      session = tracked.session;
      const { handle } = tracked;
      expect(session.state.stage).toBe('ready');
      await session.submit(command(session, 1, 'patchRangeStyle', {
        args: [0, 'A1', { bold: true, fontSize: 18 }],
      }));
      const version = handle.version();
      const history = handle.historyState();
      const captured = await session.submit(command(session, 2, 'applyOps', {
        toolbar: { id: 'paintFormat', args: null, target: selected }, selection: selected,
      }));
      expect(captured).toMatchObject({ status: 'applied', committed: false });
      const { formatToken } = captured.result as WorkbookFormatTokenResult;
      expect(typeof formatToken).toBe('string');
      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      const target = { ...selected, range: { kind: 'a1' as const, a1: 'B2' } };
      expect(await session.submit(command(session, 3, 'applyFormat', {
        target, formatToken, selection: target,
      }))).toMatchObject({ status: 'applied', committed: true });
      expect(handle.selectionFormatting(0, 'B2')).toMatchObject({ bold: true, fontSize: 18 });
      expect(session.failure).toBeUndefined();
    } finally {
      if (descriptor) Object.defineProperty(globalThis.crypto, 'randomUUID', descriptor);
      else Reflect.deleteProperty(globalThis.crypto, 'randomUUID');
      await session?.dispose();
    }
  });

  test('capture apply capture apply formats both destinations', async () => {
    const { session, handle } = await trackedSession();
    try {
      await session.submit(command(session, 1, 'patchRangeStyle', {
        args: [0, 'A2', { bold: true, fontSize: 18 }],
      }));
      await session.submit(command(session, 2, 'patchRangeStyle', {
        args: [0, 'C2', { italic: true, fontSize: 24 }],
      }));
      let seq = 2;
      const tokens: string[] = [];
      for (const [source, destination, format] of [
        ['A2', 'B2', { bold: true, fontSize: 18 }],
        ['C2', 'D2', { italic: true, fontSize: 24 }],
      ] as const) {
        const target = { ...selected, range: { kind: 'a1' as const, a1: source } };
        const captured = await session.submit(command(session, ++seq, 'applyOps', {
          toolbar: { id: 'paintFormat', args: null, target }, selection: target,
        }));
        expect(captured).toMatchObject({ status: 'applied', committed: false });
        const { formatToken } = captured.result as WorkbookFormatTokenResult;
        expect(typeof formatToken).toBe('string');
        tokens.push(formatToken);
        const into = { ...selected, range: { kind: 'a1' as const, a1: destination } };
        expect(await session.submit(command(session, ++seq, 'applyFormat', {
          target: into, formatToken, selection: into,
        }))).toMatchObject({ status: 'applied', committed: true });
        expect(handle.selectionFormatting(0, destination)).toMatchObject(format);
      }
      expect(new Set(tokens).size).toBe(2);
      expect(handle.selectionFormatting(0, 'B2')).toMatchObject({ bold: true, fontSize: 18 });
      expect(handle.selectionFormatting(0, 'D2')).toMatchObject({ italic: true, fontSize: 24 });
    } finally { await session.dispose(); }
  });

  test('capture capture without apply returns fresh formats after local cancellation', async () => {
    const { session, handle } = await trackedSession();
    try {
      await session.submit(command(session, 1, 'patchRangeStyle', { args: [0, 'A2', { bold: true }] }));
      await session.submit(command(session, 2, 'patchRangeStyle', { args: [0, 'C2', { italic: true }] }));
      const version = handle.version();
      const history = handle.historyState();
      const tokens: string[] = [];
      let seq = 2;
      for (const a1 of ['A2', 'C2']) {
        const target = { ...selected, range: { kind: 'a1' as const, a1 } };
        const captured = await session.submit(command(session, ++seq, 'applyOps', {
          toolbar: { id: 'paintFormat', args: null, target }, selection: target,
        }));
        expect(captured).toMatchObject({ status: 'applied', committed: false });
        expect(typeof (captured.result as WorkbookFormatTokenResult).formatToken).toBe('string');
        tokens.push((captured.result as WorkbookFormatTokenResult).formatToken);
        expect(handle.version()).toBe(version);
        expect(handle.historyState()).toEqual(history);
      }
      expect(new Set(tokens).size).toBe(2);
      const target = { ...selected, range: { kind: 'a1' as const, a1: 'D2' } };
      expect(await session.submit(command(session, ++seq, 'applyFormat', {
        target, formatToken: tokens[1], selection: target,
      }))).toMatchObject({ status: 'applied', committed: true });
      expect(handle.selectionFormatting(0, 'D2')).toMatchObject({ italic: true });
    } finally { await session.dispose(); }
  });

  test('unknown format token refuses without mutation', async () => {
    const { session, handle } = await trackedSession();
    try {
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      expect(await session.submit(command(session, 1, 'applyFormat', {
        target: selected, formatToken: 'unknown', selection: selected,
      }))).toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(await session.save()).toEqual(before);
      expect(await session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '9'] })))
        .toMatchObject({ status: 'applied' });
    } finally { await session.dispose(); }
  });

  test('toolbar action refuses bad args before changing selection or history', async () => {
    const { session, handle } = await trackedSession();
    const destination = { ...selected, range: { kind: 'a1' as const, a1: 'B1' } };
    try {
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const selection = structuredClone(session.selection);
      let seq = 0;
      for (const toolbar of [
        { id: 'bold', args: { value: true } },
        { id: 'italic', args: undefined },
        { id: 'paintFormat', args: {} },
        { id: 'fontSizeStep', args: { direction: 'up' } },
        { id: 'fontSize', args: { points: 401 } },
        { id: 'fontFamily', args: { family: ' ' } },
        { id: 'textColor', args: { color: 'red' } },
        { id: 'numberFormat', args: { value: 'unknown' } },
        { id: 'merge', args: { value: 'unknown' } },
        { id: 'textWrapping', args: { value: 'unknown' } },
        { id: 'save', args: null },
      ]) {
        expect(await session.submit(command(session, ++seq, 'applyOps', {
          toolbar: { ...toolbar, target: destination }, selection: destination,
        }))).toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
        expect(handle.version()).toBe(version);
        expect(handle.historyState()).toEqual(history);
        expect(session.selection).toEqual(selection);
      }
      expect(await session.save()).toEqual(before);
    } finally { await session.dispose(); }
  });

  test('rapid merge toggles resolve against current worker merges', async () => {
    const session = await createTestWorkbookSession(bytes);
    const target = { ...selected, range: { kind: 'a1' as const, a1: 'A2:B2' } };
    try {
      const payload = { toolbar: { id: 'merge', args: null, target }, selection: target };
      const first = command(session, 1, 'applyOps', payload);
      const second = command(session, 2, 'applyOps', payload, 1);
      const merging = session.submit(first);
      const unmerging = session.submit(second);
      expect(await merging).toMatchObject({ status: 'applied', committed: true });
      expect(await unmerging).toMatchObject({ status: 'applied', committed: true });
      expect(session.editState.mergedRanges).toEqual([]);
      const saved = wasm.openWorkbook(await session.save());
      try { expect(saved.mergedRanges(0, 'A2:B2')).toEqual([]); }
      finally { saved.dispose(); }
    } finally { await session.dispose(); }
  });

  test('toolbar merge refuses excessive operations before dispatch or selection changes', async () => {
    const { session, handle } = await trackedSession();
    const applying = spyOn(handle, 'applyOps').mockImplementation(() => {
      throw new Error('Oversized merge reached the engine');
    });
    try {
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const selection = structuredClone(session.selection);
      let seq = 0;
      for (const [a1, value] of [
        ['A1:B50000', 'horizontal'], ['A1:XFD2', 'vertical'],
        ['A2:B4098', 'horizontal'], ['A2:FAO3', 'vertical'],
      ] as const) {
        const target = { ...selected, range: { kind: 'a1' as const, a1 } };
        expect(await session.submit(command(session, ++seq, 'applyOps', {
          toolbar: { id: 'merge', args: { value }, target }, selection: target,
        }))).toMatchObject({ status: 'refused', reason: 'limit-exceeded', committed: false });
        expect(applying).not.toHaveBeenCalled();
        expect(handle.version()).toBe(version);
        expect(handle.historyState()).toEqual(history);
        expect(session.selection).toEqual(selection);
      }
      expect(await session.save()).toEqual(before);
    } finally { applying.mockRestore(); await session.dispose(); }
  });

  test.each(['horizontal', 'vertical'] as const)('toolbar merge admits 4096 %s operations', (value) => {
    const start = { row: 1, col: 0 };
    const end = value === 'horizontal' ? { row: 4096, col: 1 } : { row: 2, col: 4095 };
    const target = { ...selected, range: { kind: 'rowCol' as const, start, end } };
    const op = toolbarCommands.resolveWorkbookToolbar({ id: 'merge', args: { value }, target },
      { sheet: 0, range: value === 'horizontal' ? 'A2:B4097' : 'A2:FAN3', start, end }, {}, []);
    expect(op?.method).toBe('applyOps');
    if (op?.method !== 'applyOps') throw new Error('Missing merge operations');
    expect(op.args[0]).toHaveLength(4096);
    expect(op.args[0][0]).toMatchObject({ type: 'mergeCells', sheet: 0, range: { start } });
    expect(op.args[0][4095]).toMatchObject({ type: 'mergeCells', sheet: 0, range: { end } });
  });

  test('toolbar targets retain sheet identity across queued sheet insertion', async () => {
    const session = await createTestWorkbookSession(bytes, batchRequests);
    try {
      const insert = command(session, 1, 'applyOps', { args: [[{ type: 'addSheet', index: 0, name: 'New' }]] });
      const toggle = command(session, 2, 'applyOps', {
        toolbar: { id: 'bold', args: null, target: selected }, selection: selected,
      }, 1);
      const inserting = session.submit(insert);
      const toggling = session.submit(toggle);
      expect(await inserting).toMatchObject({ status: 'applied', committed: true });
      expect(await toggling).toMatchObject({ status: 'applied', committed: true });
      const saved = wasm.openWorkbook(await session.save());
      try {
        expect(saved.selectionFormatting(1, 'A1').bold).toBe(true);
        expect(saved.selectionFormatting(0, 'A1').bold).not.toBe(true);
      } finally { saved.dispose(); }
    } finally { await session.dispose(); }
  });

  test('expired format tokens refuse while the latest captured token remains usable', async () => {
    const { session, handle } = await trackedSession();
    const source = { ...selected, range: { kind: 'a1' as const, a1: 'B1' } };
    try {
      let seq = 1;
      await session.submit(command(session, seq, 'patchRangeStyle', { args: [0, 'B1', { bold: true }] }));
      const payload = { toolbar: { id: 'paintFormat', args: null, target: source } };
      const tokens: string[] = [];
      for (let at = 0; at < 65; at++) {
        const captured = await session.submit(command(session, ++seq, 'applyOps', payload));
        expect(captured).toMatchObject({ status: 'applied', committed: false });
        tokens.push((captured.result as WorkbookFormatTokenResult).formatToken);
        if (at < 64) {
          const recaptured = await session.submit(command(session, ++seq, 'applyOps', payload));
          expect(typeof (recaptured.result as WorkbookFormatTokenResult).formatToken).toBe('string');
        }
      }
      expect(new Set(tokens).size).toBe(65);
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      expect(await session.submit(command(session, ++seq, 'applyFormat', {
        target: selected, formatToken: tokens[0],
      }))).toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(await session.save()).toEqual(before);
      expect(await session.submit(command(session, ++seq, 'applyFormat', {
        target: selected, formatToken: tokens[tokens.length - 1], selection: selected,
      }))).toMatchObject({ status: 'applied', committed: true });
      expect(session.editState.formatting?.bold).toBe(true);
    } finally { await session.dispose(); }
  });

  test('missing format tokens refuse without mutation', async () => {
    const { session, handle } = await trackedSession();
    try {
      const before = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      let seq = 0;
      for (const payload of [{ target: selected }, { target: selected, formatToken: null },
        { target: selected, formatToken: '' }]) {
        expect(await session.submit(command(session, ++seq, 'applyFormat',
          payload as WorkbookCommandPayload<'applyFormat'>)))
          .toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
      }
      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(await session.save()).toEqual(before);
    } finally { await session.dispose(); }
  });

  test('moveSheet refreshes sheet order and preserves stable edit targets', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const original = structuredClone(session.state.sheets);
      const target = { ...selected, sheetId: original[0].id };
      expect(await session.submit(command(session, 1, 'applyOps', {
        args: [[{ type: 'moveSheet', from: 0, to: 1 }]],
      }))).toMatchObject({ status: 'applied', committed: true });
      expect(session.state.sheets).toEqual([
        { ...original[1], index: 0 }, { ...original[0], index: 1 },
      ]);
      expect(await session.submit(command(session, 2, 'editCell', {
        args: [0, 1, 0, '7'], target: { ...target, range: { kind: 'a1', a1: 'A2' } },
      }))).toMatchObject({ status: 'applied', committed: true });
      expect((await session.call.cellInputs(1, 'A2')).cells[0][0].input).toBe('7');
      expect((await session.call.cellInputs(0, 'A2')).cells[0][0].input).toBe('');
      expect(await session.submit(command(session, 3, 'applyOps', { args: [[
        { type: 'moveSheet', from: 1, to: 0 },
        { type: 'setCell', sheet: 1, target, at: { row: 2, col: 0 },
          cell: { value: { kind: 'number', value: 9 }, formula: null, style: null } },
      ]] }))).toMatchObject({ status: 'applied', committed: true });
      expect(session.state.sheets).toEqual(original);
      expect((await session.call.cellInputs(0, 'A3')).cells[0][0].input).toBe('9');
      expect((await session.call.cellInputs(1, 'A3')).cells[0][0].input).toBe('');
    } finally { await session.dispose(); }
  });

  test('moveSheet resolves the source sheet from its stable target', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const original = structuredClone(session.state.sheets);
      expect(await session.submit(command(session, 1, 'applyOps', {
        args: [[{ type: 'moveSheet', from: 0, to: 0, target: { ...selected, sheetId: original[1].id } }]],
      }))).toMatchObject({ status: 'applied', committed: true });
      expect(session.state.sheets).toEqual([
        { ...original[1], index: 0 }, { ...original[0], index: 1 },
      ]);
      expect((await session.call.cellInputs(1, 'A1')).cells[0][0].input).toBe('1');
    } finally { await session.dispose(); }
  });

  test('mixed sheet batch commits stable targets and replays undo and redo', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const original = structuredClone(session.state.sheets);
      const targetA = { ...selected, sheetId: original[0].id };
      const targetB = { ...selected, sheetId: original[1].id };
      expect(await session.submit(command(session, 1, 'applyOps', { args: [[
        { type: 'addSheet', index: 2, name: 'C' },
        { type: 'moveSheet', from: 0, to: 2, target: targetA },
        { type: 'renameSheet', sheet: 0, name: 'Renamed', target: targetB },
        { type: 'setCell', sheet: 2, target: targetA, at: { row: 1, col: 0 },
          cell: { value: { kind: 'number', value: 42 }, formula: null, style: null } },
      ]] }))).toMatchObject({ status: 'applied', committed: true });
      const after = structuredClone(session.state.sheets);
      expect(after.map((sheet) => sheet.name)).toEqual(['Renamed', 'C', original[0].name]);
      expect(after[0].id).toBe(original[1].id);
      expect(after[2].id).toBe(original[0].id);
      expect((await session.call.cellInputs(2, 'A2')).cells[0][0].input).toBe('42');
      expect(await session.submit(command(session, 2, 'undo', { args: [] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect(session.state.sheets).toEqual(original);
      expect((await session.call.cellInputs(0, 'A2')).cells[0][0].input).toBe('');
      expect(await session.submit(command(session, 3, 'redo', { args: [] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect(session.state.sheets).toEqual(after);
      expect((await session.call.cellInputs(2, 'A2')).cells[0][0].input).toBe('42');
    } finally { await session.dispose(); }
  });

  for (const structural of ['move', 'insert', 'removeBefore'] as const) {
    test(`active sheet follows ${structural} and rename through undo and redo`, async () => {
      const { session, handle } = await trackedSession();
      try {
        const original = structuredClone(session.state.sheets);
        const before = structural === 'removeBefore' ? 1 : 0;
        const after = structural === 'removeBefore' ? 0 : 1;
        const target = { ...selected, sheetId: original[before].id };
        await session.submit(command(session, 1, 'setActiveSheet', { args: [before] }));
        const operation = structural === 'move' ? { type: 'moveSheet', from: 0, to: 1 }
          : structural === 'insert' ? { type: 'addSheet', index: 0, name: 'C' }
          : { type: 'removeSheet', index: 0 };
        expect(await session.submit(command(session, 2, 'applyOps', { args: [[
          operation, { type: 'renameSheet', sheet: after, name: 'Renamed', target },
        ]] }))).toMatchObject({ status: 'applied', committed: true });
        expect(session.state.activeSheet).toBe(after);
        expect(session.state.sheets[session.state.activeSheet].id).toBe(target.sheetId);
        expect(handle.sheetInfo().activeSheet).toBe(after);
        expect(await session.submit(command(session, 3, 'undo', { args: [] })))
          .toMatchObject({ status: 'applied', committed: true });
        expect(session.state.activeSheet).toBe(before);
        expect(session.state.sheets).toEqual(original);
        expect(await session.submit(command(session, 4, 'redo', { args: [] })))
          .toMatchObject({ status: 'applied', committed: true });
        expect(session.state.activeSheet).toBe(after);
        expect(session.state.sheets[after]).toMatchObject({ id: target.sheetId, name: 'Renamed' });
        expect(session.failure).toBeUndefined();
      } finally { await session.dispose(); }
    });
  }

  test('save preserves workbook view tabs through move undo and redo', async () => {
    const original = await JSZip.loadAsync(bytes);
    const sourceXml = await original.file('xl/workbook.xml')!.async('string');
    original.file('xl/workbook.xml', sourceXml
      .replace('<sheets>', '<bookViews><workbookView/></bookViews><sheets>')
      .replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>'));
    const session = await createTestWorkbookSession(await original.generateAsync({ type: 'uint8array' }));
    try {
      expect(await session.submit(command(session, 1, 'applyOps', { args: [[
        { type: 'renameSheet', sheet: 0, name: 'Renamed' },
        { type: 'moveSheet', from: 1, to: 0 },
      ]] }))).toMatchObject({ status: 'applied', committed: true });
      const zip = await JSZip.loadAsync(await session.save());
      const xml = await zip.file('xl/workbook.xml')!.async('string');
      expect(xml).toContain('<workbookView activeTab="1" firstSheet="1"/>');
      expect(session.state.activeSheet).toBe(1);
      await session.submit(command(session, 2, 'undo', { args: [] }));
      const undone = await JSZip.loadAsync(await session.save());
      expect(await undone.file('xl/workbook.xml')!.async('string'))
        .toBe(await original.file('xl/workbook.xml')!.async('string'));
      await session.submit(command(session, 3, 'redo', { args: [] }));
      const redone = await JSZip.loadAsync(await session.save());
      expect(await redone.file('xl/workbook.xml')!.async('string')).toBe(xml);
    } finally { await session.dispose(); }
  });

  test('queued numeric moveSheet refuses after concurrent sheet deletion', async () => {
    const session = await createTestWorkbookSession(bytes, batchRequests);
    try {
      const remove = command(session, 1, 'applyOps', { args: [[{ type: 'removeSheet', index: 1 }]] });
      const move = command(session, 2, 'applyOps', { args: [[{ type: 'moveSheet', from: 0, to: 0 }]] }, 1);
      const statuses = await Promise.all([session.submit(remove), session.submit(move)]);
      expect(statuses[0]).toMatchObject({ status: 'applied', committed: true });
      expect(statuses[1]).toMatchObject({ status: 'refused', reason: 'stale-target', committed: false });
      expect(session.state.sheets).toEqual([{ id: 'sheet:0', index: 0, name: 'Data' }]);
      const saved = await session.save();
      await session.submit(command(session, 3, 'undo', { args: [] }));
      expect(session.state.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Other']);
      await session.submit(command(session, 4, 'redo', { args: [] }));
      expect(await session.save()).toEqual(saved);
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('moveSheet undo and redo restore sheet ids and later targeted edits', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const original = structuredClone(session.state.sheets);
      await session.submit(command(session, 1, 'applyOps', { args: [[{ type: 'moveSheet', from: 0, to: 1 }]] }));
      const moved = [{ ...original[1], index: 0 }, { ...original[0], index: 1 }];
      expect(session.state.sheets).toEqual(moved);
      expect(await session.submit(command(session, 2, 'undo', { args: [] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect(session.state.sheets).toEqual(original);
      expect(await session.submit(command(session, 3, 'redo', { args: [] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect(session.state.sheets).toEqual(moved);
      await session.submit(command(session, 4, 'undo', { args: [] }));
      expect(session.state.sheets).toEqual(original);
      expect(await session.submit(command(session, 5, 'editCell', {
        args: [1, 1, 0, '8'], target: { sheetId: original[0].id, range: { kind: 'a1', a1: 'A2' } },
      }))).toMatchObject({ status: 'applied', committed: true });
      expect((await session.call.cellInputs(0, 'A2')).cells[0][0].input).toBe('8');
      expect((await session.call.cellInputs(1, 'A2')).cells[0][0].input).toBe('');
    } finally { await session.dispose(); }
  });

  for (const kind of ['applyOps', 'propose'] as const) {
    test(`queued ${kind} resolves individual stable sheet targets after insertion`, async () => {
      const session = await createTestWorkbookSession(bytes, batchRequests);
      try {
        const insert = command(session, 1, 'applyOps', { args: [[{ type: 'addSheet', index: 0, name: 'New' }]] });
        const payload = kind === 'applyOps'
          ? { args: [[{ type: 'setCell', sheet: 0, target: selected, at: { row: 1, col: 0 },
            cell: { value: { kind: 'number', value: 7 }, formula: null, style: null } }]] }
          : { args: ['agent', null, [{ sheet: 0, target: selected, row: 1, col: 0, input: '7' }]] };
        const anchored = command(session, 2, kind, payload as WorkbookCommandPayload<typeof kind>, 1);
        const numeric = command(session, 3, 'editCell', { args: [0, 1, 0, 'wrong'] }, 1);
        const pending = [session.submit(insert), session.submit(anchored), session.submit(numeric)];
        const statuses = await Promise.all(pending);
        expect(statuses.map((status) => status.status)).toEqual(['applied', 'applied', 'refused']);
        expect(statuses[2].reason).toBe('stale-target');
        if (kind === 'propose') {
          expect((statuses[1].result as wasm.Proposal).cells[0].sheet).toBe(1);
          await session.submit(command(session, 4, 'acceptProposal', { args: [(statuses[1].result as wasm.Proposal).id] }));
        }
        expect((await session.call.cellInputs(0, 'A2')).cells[0][0].input).toBe('');
        expect((await session.call.cellInputs(1, 'A2')).cells[0][0].input).toBe('7');
        expect(session.failure).toBeUndefined();
      } finally { await session.dispose(); }
    });

    test(`queued numeric ${kind} refuses after sheet deletion`, async () => {
      const session = await createTestWorkbookSession(bytes, batchRequests);
      try {
        const remove = command(session, 1, 'applyOps', { args: [[{ type: 'removeSheet', index: 1 }]] });
        const payload = kind === 'applyOps'
          ? { args: [[{ type: 'renameSheet', sheet: 0, name: 'Wrong' }]] }
          : { args: ['agent', null, [{ sheet: 0, row: 1, col: 0, input: 'wrong' }]] };
        const edit = command(session, 2, kind, payload as WorkbookCommandPayload<typeof kind>, 1);
        const statuses = await Promise.all([session.submit(remove), session.submit(edit)]);
        expect(statuses[1]).toMatchObject({ status: 'refused', reason: 'stale-target', committed: false });
        expect(session.state.sheets[0].name).toBe('Data');
        expect(session.failure).toBeUndefined();
      } finally { await session.dispose(); }
    });
  }

  test('queued stable targets survive sheet insertion and refuse deleted sheets', async () => {
    const session = await createTestWorkbookSession(bytes, batchRequests);
    try {
      const original = { ...selected, sheetId: session.state.sheets[0].id };
      const insert = command(session, 1, 'applyOps', { args: [[{ type: 'addSheet', index: 0, name: 'New' }]] });
      const edit = command(session, 2, 'editCell', { args: [0, 0, 0, '7'], target: original }, 1);
      expect((await Promise.all([session.submit(insert), session.submit(edit)])).map((status) => status.status))
        .toEqual(['applied', 'applied']);
      expect(session.state.sheets.map((sheet) => sheet.id)).toEqual([
        expect.any(String), original.sheetId, 'sheet:1',
      ]);
      expect(session.state.sheets[0].id).not.toBe(original.sheetId);
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('');
      expect((await session.call.cellInputs(1, 'A1')).cells[0][0].input).toBe('7');
      expect(session.selection?.sheetId).toBe(original.sheetId);
      await session.submit(command(session, 3, 'undo', { args: [] }));
      await session.submit(command(session, 4, 'undo', { args: [] }));
      expect(session.state.sheets.map((sheet) => sheet.id)).toEqual(['sheet:0', 'sheet:1']);
      await session.submit(command(session, 5, 'redo', { args: [] }));
      const remove = command(session, 6, 'applyOps', { args: [[{ type: 'removeSheet', index: 1 }]] });
      const deleted = command(session, 7, 'editCell', { args: [0, 0, 0, 'wrong'], target: original }, 6);
      await session.submit(remove);
      const saved = await session.save();
      expect(await session.submit(deleted)).toMatchObject({
        status: 'refused', reason: 'target-sheet-missing', committed: false,
      });
      expect(await session.save()).toEqual(saved);
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('');
      await session.submit(command(session, 8, 'undo', { args: [] }));
      expect(await session.submit(command(session, 9, 'editCell', { args: [0, 0, 0, '8'], target: original })))
        .toMatchObject({ status: 'applied' });
      expect((await session.call.cellInputs(1, 'A1')).cells[0][0].input).toBe('8');
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('ordered catalogs reads and batches use acknowledged sheet ids', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      await session.submit(command(session, 1, 'applyOps', { args: [[{ type: 'addSheet', index: 0, name: 'New' }]] }));
      const added = { ...selected, sheetId: session.state.sheets[0].id };
      expect(await session.call.sheets()).toEqual(session.state.sheets);
      const read = await session.call.readCells({ ranges: [selected, added] });
      if (!read.ok) throw new Error('Expected cell read');
      expect(read.sheets.map((sheet) => sheet.sheetId)).toEqual(session.state.sheets.map((sheet) => sheet.id));
      expect(read.ranges.map((range) => range.target)).toEqual([selected, added]);
      expect(read.ranges.map((range) => range.cells[0][0].displayText)).toEqual(['1', '']);
      const request = { expectVersion: session.confirmedBasis.version, steps: [
        { op: 'setCellInputs' as const, target: selected, inputs: [['7']] },
        { op: 'setCellInputs' as const, target: added, inputs: [['9']] },
      ] };
      const validation = await session.call.validateEdits(request);
      if (!validation.ok) throw new Error('Expected validation');
      expect(validation.previews.map((preview) => preview.target)).toEqual([selected, added]);
      const applied = await session.submit(command(session, 2, 'applyEdits', { args: [request] }));
      const result = applied.result as XlsxEditResult;
      if (!result.ok) throw new Error('Expected batch result');
      expect(result.receipts.map((receipt) => receipt.target)).toEqual([selected, added]);
      expect(result.receipts.map((receipt) => receipt.changedCells[0].sheetId)).toEqual([selected.sheetId, added.sheetId]);
      expect(result.changedSheets).toEqual([added.sheetId, selected.sheetId]);
      expect(result.calculation.changed).toContainEqual({ sheetId: selected.sheetId, row: 0, col: 1, a1: 'B1' });
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('9');
      expect((await session.call.cellInputs(1, 'A1')).cells[0][0].input).toBe('7');
      const found = await session.call.findText({ text: '14', sheetIds: [selected.sheetId] });
      if (!found.ok) throw new Error('Expected text search');
      expect(found.matches[0].cell.sheetId).toBe(selected.sheetId);
      expect(await session.submit(command(session, 3, 'applyEdits', { args: [request] })))
        .toMatchObject({ status: 'refused', reason: 'stale-version', committed: false });
      expect(await session.call.validateEdits(request)).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      const locked = { ...selected, range: { kind: 'a1' as const, a1: 'E1' } };
      expect(await session.submit(command(session, 4, 'applyEdits', { args: [{
        expectVersion: session.confirmedBasis.version, steps: [{ op: 'setCellInputs', target: locked, inputs: [['bad']] }],
      }] }))).toMatchObject({ status: 'refused', reason: 'locked-target', result: { failure: { target: locked } } });
      await session.submit(command(session, 5, 'undo', { args: [] }));
      expect(await session.call.sheets()).toEqual(session.state.sheets);
      await session.submit(command(session, 6, 'undo', { args: [] }));
      expect(await session.call.readCells({ ranges: [added] })).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
      expect(await session.submit(command(session, 7, 'applyEdits', { args: [{
        expectVersion: session.confirmedBasis.version, steps: [{ op: 'setCellInputs', target: added, inputs: [['bad']] }],
      }] }))).toMatchObject({ status: 'refused', reason: 'target-sheet-missing', committed: false });
      await session.submit(command(session, 8, 'redo', { args: [] }));
      const restored = await session.call.readCells({ ranges: [added] });
      expect(restored).toMatchObject({ ok: true, ranges: [{ target: added }] });
      expect(await session.call.sheets()).toEqual(session.state.sheets);
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  for (const operation of [
    { type: 'setColWidth', sheet: 0, col: 0, width: -1 },
    { type: 'setColWidth', sheet: 0, col: 0, width: 256 },
    { type: 'setRowHeight', sheet: 0, row: 0, height: -1 },
    { type: 'setRowHeight', sheet: 0, row: 0, height: 410 },
  ]) {
    test(`refuses invalid ${operation.type} dimensions before dispatch (${operation.width ?? operation.height})`, async () => {
      const { session, handle } = await trackedSession();
      const applying = spyOn(handle, 'applyOps');
      try {
        const saved = await session.save();
        const version = session.confirmedBasis.version;
        const history = handle.historyState();
        expect(await session.submit(command(session, 1, 'applyOps', { args: [[
          { type: 'renameSheet', sheet: 0, name: 'Uncommitted' }, operation,
        ]] }))).toMatchObject({ status: 'refused', reason: 'invalid-step', committed: false });
        expect(applying).not.toHaveBeenCalled();
        expect(await session.save()).toEqual(saved);
        expect(session.confirmedBasis.version).toBe(version);
        expect(handle.historyState()).toEqual(history);
        expect(await session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '2'] })))
          .toMatchObject({ status: 'applied' });
        expect(session.failure).toBeUndefined();
      } finally { applying.mockRestore(); await session.dispose(); }
    });
  }

  for (const kind of ['chartsheet', 'dialogsheet']) {
    test(`navigates to a non-editable ${kind} without worksheet formatting`, async () => {
      const zip = await JSZip.loadAsync(bytes);
      const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
      zip.file('[Content_Types].xml', contentTypes.replace(
        '/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
        `/xl/${kind}s/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${kind}+xml`
      ));
      const relationships = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
      zip.file('xl/_rels/workbook.xml.rels', relationships.replace(
        '/worksheet" Target="worksheets/sheet2.xml', `/${kind}" Target="${kind}s/sheet2.xml`
      ));
      zip.file(`xl/${kind}s/sheet2.xml`, `<${kind} xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`);
      const session = await createTestWorkbookSession(await zip.generateAsync({ type: 'uint8array' }));
      try {
        expect(await session.submit(command(session, 1, 'setActiveSheet', { args: [1] })))
          .toMatchObject({ status: 'applied', committed: false });
        expect(session.editState).toMatchObject({ activeSheet: 1, selectionEditable: false,
          formattingAvailable: false, formatting: {} });
        expect(await session.submit(command(session, 2, 'editCell', { args: [1, 0, 0, 'bad'] })))
          .toMatchObject({ status: 'refused', committed: false });
        const saved = await session.save();
        const version = session.confirmedBasis.version;
        expect(await session.submit(command(session, 3, 'applyFormat', {
          target: selected, capturedFrom: { ...selected, sheetId: 'sheet:1' },
        }))).toMatchObject({ status: 'refused', reason: 'unsupported', committed: false });
        expect(await session.save()).toEqual(saved);
        expect(session.confirmedBasis.version).toBe(version);
        await session.submit(command(session, 4, 'setActiveSheet', { args: [0] }));
        expect(session.editState).toMatchObject({ selectionEditable: true, formattingAvailable: true });
        expect(session.failure).toBeUndefined();
      } finally { await session.dispose(); }
    });
  }

  for (const calculation of [undefined, { nowSerial: 46_000.5, randSeed: 123456789 }]) {
    test(`safe refusals restore the ${calculation ? 'preceding' : 'absent'} calculation context`, async () => {
      const { session, handle } = await trackedSession(calculation);
      const contexts = spyOn(handle, 'setCalculationContext');
      const clock = spyOn(Date, 'now').mockReturnValue((47_000.75 - 25_569) * 86_400_000);
      const main = wasm.openWorkbook(bytes, { calculation });
      try {
        const override = { nowSerial: 50_000.25, randSeed: 987654321 };
        expect(await session.submit(command(session, 1, 'applyOps', {
          args: [[{ type: 'notAnOperation' }]], calculation: override,
        }))).toMatchObject({ status: 'refused', committed: false });
        expect(await session.submit(command(session, 2, 'applyEdits', {
          args: [{ expectVersion: session.confirmedBasis.version,
            steps: [{ op: 'setCellInputs', target: { ...selected, range: { kind: 'a1', a1: 'E1' } }, inputs: [['locked']] }] }],
          calculation: override,
        }))).toMatchObject({ status: 'refused', reason: 'locked-target', committed: false });
        expect(contexts.mock.calls).toEqual([[override], [calculation ?? null], [override], [calculation ?? null]]);
        const input = calculation ? '=NOW()+TODAY()+RANDBETWEEN(1,1000000)' : '=NOW()+TODAY()';
        await session.submit(command(session, 3, 'editCell', { args: [0, 1, 0, input] }));
        main.editCell(0, 1, 0, input);
        const request = { ranges: [{ ...selected, range: { kind: 'a1' as const, a1: 'A2' } }] };
        const expected = main.readCells(request);
        const actual = await session.call.readCells(request);
        if (!expected.ok || !actual.ok) throw new Error('Expected cell read');
        expect(actual.ranges).toEqual(expected.ranges);
        expect(actual.ranges[0].cells[0][0].value.kind).toBe('number');
        if (!calculation) expect(actual.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: 94_000.75 });
        await session.submit(command(session, 4, 'editCell', { args: [0, 4, 0, '=NOW()'], calculation: override }));
        main.setCalculationContext(override);
        main.editCell(0, 4, 0, '=NOW()');
        const overridden = await session.call.readCells({ ranges: [{ ...selected, range: { kind: 'a1', a1: 'A5' } }] });
        if (!overridden.ok) throw new Error('Expected cell read');
        expect(overridden.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: override.nowSerial });
        expect(await session.submit(command(session, 5, 'applyOps', {
          args: [[{ type: 'notAnOperation' }]], calculation: { nowSerial: 1, randSeed: 1 },
        }))).toMatchObject({ status: 'refused' });
        await session.submit(command(session, 6, 'editCell', { args: [0, 2, 0, '=NOW()+RANDBETWEEN(1,1000000)'] }));
        main.editCell(0, 2, 0, '=NOW()+RANDBETWEEN(1,1000000)');
        const latest = { ranges: [{ ...selected, range: { kind: 'a1' as const, a1: 'A3' } }] };
        const latestExpected = main.readCells(latest);
        const latestActual = await session.call.readCells(latest);
        if (!latestExpected.ok || !latestActual.ok) throw new Error('Expected cell read');
        expect(latestActual.ranges).toEqual(latestExpected.ranges);
        await session.submit(command(session, 7, 'editCell', { args: [0, 3, 0, '=TODAY()'] }));
        const today = await session.call.readCells({ ranges: [{ ...selected, range: { kind: 'a1', a1: 'A4' } }] });
        if (!today.ok) throw new Error('Expected cell read');
        expect(today.ranges[0].cells[0][0].value).toEqual({ kind: 'number', value: Math.floor(override.nowSerial) });
        expect(session.failure).toBeUndefined();
      } finally { clock.mockRestore(); contexts.mockRestore(); main.dispose(); await session.dispose(); }
    });
  }

  test('checks Enter locks without staging and validates a rectangular paste once', async () => {
    const { session, handle } = await trackedSession();
    const validating = spyOn(handle, 'validateEdits');
    const batches = spyOn(handle, 'applyEdits');
    const entering = spyOn(handle, 'editCell');
    try {
      await session.submit(command(session, 1, 'editCell', { args: [0, 0, 0, '3'] }));
      expect(entering).toHaveBeenCalledWith(0, 0, 0, '3');
      expect(batches).not.toHaveBeenCalled();
      expect(validating.mock.calls[0][0].steps).toEqual([{ op: 'patchStyle',
        target: { ...selected, range: { kind: 'rowCol', start: { row: 0, col: 0 }, end: { row: 0, col: 0 } } }, patch: {} }]);
      validating.mockClear();
      const edits = Array.from({ length: 200 }, (_, row) => Array.from({ length: 3 }, (_, col) =>
        ({ row: row + 1, col, input: String(row * 3 + col) }))).flat();
      expect(await session.submit(command(session, 2, 'editCells', { args: [0, edits] })))
        .toMatchObject({ status: 'applied' });
      expect(validating).toHaveBeenCalledTimes(1);
      expect(validating.mock.calls[0][0].steps).toHaveLength(1);
      expect(validating.mock.calls[0][0].steps[0]).toMatchObject({ op: 'patchStyle', patch: {} });
      expect((await session.call.cellInputs(0, 'A2:C201')).cells.map((row) => row.map((cell) => cell.input)))
        .toEqual(Array.from({ length: 200 }, (_, row) => Array.from({ length: 3 }, (_, col) => String(row * 3 + col))));
      expect(handle.historyState().undoDepth).toBe(2);
      await session.submit(command(session, 3, 'undo', { args: [] }));
      expect(handle.historyState().undoDepth).toBe(1);
      expect((await session.call.cellInputs(0, 'A2:C201')).cells.flat().every((cell) => cell.input === '')).toBe(true);
      expect((await session.call.cellInputs(0, 'A2')).cells[0][0].input).toBe('');
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('3');
    } finally { entering.mockRestore(); batches.mockRestore(); validating.mockRestore(); await session.dispose(); }
  });

  for (const kind of ['editCells', 'propose'] as const) {
    test(`validates sparse ${kind} in bounded chunks and commits one undo step`, async () => {
      const { session, handle } = await trackedSession();
      const validating = spyOn(handle, 'validateEdits');
      const batches = spyOn(handle, 'applyEdits');
      const pasting = spyOn(handle, 'editCells');
      const proposing = spyOn(handle, 'propose');
      try {
        const edits = Array.from({ length: 200 }, (_, index) => ({ row: index * 2 + 1, col: index % 3, input: String(index) }));
        const status = kind === 'editCells'
          ? await session.submit(command(session, 1, kind, { args: [0, edits] }))
          : await session.submit(command(session, 1, kind, { args: ['agent', null, edits.map((edit) => ({ ...edit, sheet: 0 }))] }));
        expect(status.status).toBe('applied');
        expect(validating.mock.calls.map(([request]) => request.steps.length)).toEqual([128, 72]);
        expect(validating.mock.calls.every(([request]) => request.steps.every((step) =>
          step.op === 'patchStyle' && Object.keys(step.patch).length === 0))).toBe(true);
        expect(batches).not.toHaveBeenCalled();
        expect(kind === 'editCells' ? pasting : proposing).toHaveBeenCalledTimes(1);
        if (kind === 'propose') {
          expect(handle.historyState().undoDepth).toBe(0);
          expect(await session.submit(command(session, 2, 'acceptProposal', { args: [(status.result as wasm.Proposal).id] })))
            .toMatchObject({ status: 'applied', committed: true });
        }
        expect(handle.historyState().undoDepth).toBe(1);
        expect((await session.call.cellInputs(0, 'B4')).cells[0][0].input).toBe('1');
        await session.submit(command(session, kind === 'propose' ? 3 : 2, 'undo', { args: [] }));
        expect(handle.historyState().undoDepth).toBe(0);
        expect((await session.call.cellInputs(0, 'A2:C400')).cells.flat().every((cell) => cell.input === '')).toBe(true);
      } finally { proposing.mockRestore(); pasting.mockRestore(); batches.mockRestore(); validating.mockRestore(); await session.dispose(); }
    });
  }

  test('large paste refusals leave version content and history unchanged', async () => {
    const { session, handle } = await trackedSession();
    const pasting = spyOn(handle, 'editCells');
    try {
      await session.submit(command(session, 1, 'applyOps', { args: [[
        { type: 'unmergeCells', sheet: 0, range: { start: { row: 0, col: 3 }, end: { row: 0, col: 4 } } },
        { type: 'mergeCells', sheet: 0, range: { start: { row: 401, col: 3 }, end: { row: 401, col: 4 } } },
      ]] }));
      await session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '3'] }));
      await session.submit(command(session, 3, 'undo', { args: [] }));
      const saved = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const edits = Array.from({ length: 200 }, (_, index) => ({ row: index * 2 + 1, col: 0, input: String(index) }));
      let seq = 3;
      for (const [sheet, extra, reason] of [
        [0, { row: 401, col: 4, input: 'locked' }, 'locked-target'],
        [0, { row: 401, col: 0, input: 'x'.repeat(32_768) }, 'invalid-step'],
        [99, { row: 401, col: 0, input: 'missing' }, 'missing-target'],
      ] as const) {
        expect(await session.submit(command(session, ++seq, 'editCells', { args: [sheet, [...edits, extra]] })))
          .toMatchObject({ status: 'refused', reason, committed: false });
        expect(await session.save()).toEqual(saved);
        expect(handle.version()).toBe(version);
        expect(handle.historyState()).toEqual(history);
      }
      expect(pasting).toHaveBeenCalledTimes(1);
      expect(session.failure).toBeUndefined();
    } finally { pasting.mockRestore(); await session.dispose(); }
  });

  for (const budget of ['cells', 'text'] as const) {
    test(`splits paste admission rectangles at the ${budget} budget without staging`, async () => {
      const { session, handle } = await trackedSession();
      const validating = spyOn(handle, 'validateEdits');
      const pasting = spyOn(handle, 'editCells');
      const batches = spyOn(handle, 'applyEdits');
      try {
        const edits = budget === 'cells'
          ? Array.from({ length: 100_005 }, (_, at) => ({ row: Math.floor(at / 5) + 1, col: at % 5, input: '7' }))
          : Array.from({ length: 40 }, (_, col) => ({ row: 1, col, input: 'x'.repeat(32_767) }));
        expect(await session.submit(command(session, 1, 'editCells', { args: [0, edits] })))
          .toMatchObject({ status: 'applied', committed: true });
        expect(validating).toHaveBeenCalledTimes(2);
        for (const [request] of validating.mock.calls) {
          expect(request.steps.length).toBeLessThanOrEqual(128);
          let cells = 0;
          for (const step of request.steps) {
            expect(step).toMatchObject({ op: 'patchStyle', patch: {} });
            if (step.target.range.kind !== 'rowCol') throw new Error('Expected rectangle');
            const { start, end } = step.target.range;
            cells += (end.row - start.row + 1) * (end.col - start.col + 1);
          }
          expect(cells).toBeLessThanOrEqual(100_000);
          if (budget === 'text') expect(cells * 32_767).toBeLessThanOrEqual(1_048_576);
          expect(new TextEncoder().encode(JSON.stringify(request)).length).toBeLessThanOrEqual(16 * 1024 * 1024);
        }
        expect(pasting).toHaveBeenCalledTimes(1);
        expect(batches).not.toHaveBeenCalled();
        expect(handle.historyState().undoDepth).toBe(1);
        const last = edits[edits.length - 1];
        expect(handle.cell(0, last.row, last.col).input).toBe(last.input);
        await session.submit(command(session, 2, 'undo', { args: [] }));
        expect(handle.cell(0, last.row, last.col).input).toBe('');
        expect(handle.cell(0, 1, 0).input).toBe('');
        expect(session.failure).toBeUndefined();
      } finally { batches.mockRestore(); pasting.mockRestore(); validating.mockRestore(); await session.dispose(); }
    });
  }

  test('single-cell receipts retain recalculated and limited addresses on other sheets', async () => {
    const session = await createTestWorkbookSession(bytes);
    const main = wasm.openWorkbook(bytes);
    try {
      const setup: [number, number, number, string][] = [
        [0, 1_048_575, 16_383, '1'],
        [1, 0, 0, '=Data!A1*3'],
        [1, 0, 1, '=SUM(Data!A1:XFD1048576)'],
      ];
      let seq = 0;
      for (const args of setup) {
        main.editCell(...args);
        await session.submit(command(session, ++seq, 'editCell', { args }));
      }
      const status = await session.submit(command(session, ++seq, 'editCell', { args: [0, 0, 0, '7'] }));
      const expected = main.editCell(0, 0, 0, '7');
      const actual = status.result as wasm.EditResult;
      expect(actual.changed ?? []).toEqual(expected.changed ?? []);
      expect(actual.changed ?? []).toContain('B1');
      expect(actual.changed ?? []).toContain('Other!A1');
      expect(actual.limitedCells ?? []).toEqual(expected.limitedCells ?? []);
      expect(actual.limitedCells ?? []).toContain('Other!B1');
      main.setActiveSheet(1);
      await session.submit(command(session, ++seq, 'setActiveSheet', { args: [1] }));
      const inactive = await session.submit(command(session, ++seq, 'editCell', { args: [0, 0, 0, '8'] }));
      const inactiveExpected = main.editCell(0, 0, 0, '8');
      const inactiveActual = inactive.result as wasm.EditResult;
      expect(inactiveActual.changed ?? []).toEqual(inactiveExpected.changed ?? []);
      expect(inactiveActual.changed ?? []).toContain('Data!B1');
      expect(inactiveActual.changed ?? []).toContain('A1');
      expect(inactiveActual.limitedCells ?? []).toEqual(inactiveExpected.limitedCells ?? []);
      expect(inactiveActual.limitedCells ?? []).toContain('B1');
    } finally { main.dispose(); await session.dispose(); }
  });

  test('large selections report unavailable formatting and retain all intersecting merges', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      await session.submit(command(session, 1, 'patchRangeStyle', { args: [0, 'A1', { bold: true }] }));
      expect(session.editState.formatting?.bold).toBe(true);
      const selection = { ...selected, range: { kind: 'a1' as const, a1: 'A1:E20001' }, activeCell: { row: 0, col: 0 } };
      await session.submit(command(session, 2, 'select', { selection }));
      expect(session.editState).toMatchObject({ formattingAvailable: false, formatting: {},
        mergedRanges: [{ start: { row: 0, col: 3 }, end: { row: 0, col: 4 } }] });
      await session.submit(command(session, 3, 'select', { selection: { ...selected, range: { kind: 'a1', a1: 'A1:B1' } } }));
      expect(session.editState.formattingAvailable).toBe(true);
      expect(session.editState.formatting?.bold).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('sheet navigation preserves explicit selections and same-sheet noops', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const first = { ...selected, range: { kind: 'a1' as const, a1: 'B2:C3' }, activeCell: { row: 2, col: 2 } };
      await session.submit(command(session, 1, 'select', { selection: first }));
      expect(await session.submit(command(session, 2, 'setActiveSheet', { args: [0] })))
        .toMatchObject({ status: 'noop', committed: false });
      expect(session.selection).toEqual(first);
      const destination = { ...first, sheetId: 'sheet:1' };
      expect(await session.submit(command(session, 3, 'setActiveSheet', { args: [1], selection: destination })))
        .toMatchObject({ status: 'applied', committed: false });
      expect(session.selection).toEqual(destination);
      expect(await session.submit(command(session, 4, 'setActiveSheet', { args: [1] })))
        .toMatchObject({ status: 'noop', committed: false });
      expect(session.selection).toEqual(destination);
      const replacement = { ...destination, range: { kind: 'a1' as const, a1: 'D4' }, activeCell: { row: 3, col: 3 } };
      expect(await session.submit(command(session, 5, 'setActiveSheet', { args: [1], selection: replacement })))
        .toMatchObject({ status: 'applied', committed: false });
      expect(session.selection).toEqual(replacement);
    } finally { await session.dispose(); }
  });

  test('pipelines anchored commands on confirmed basis', async () => {
    const sent: WorkbookCommand[] = [];
    const session = await createTestWorkbookSession(bytes, (transport) => {
      const batched = batchRequests(transport);
      return { ...batched, listen(listener) {
        return batched.listen((message) => {
          const input = capturedWorkbookCommand(message);
          if (input) sent.push(input);
          listener(message);
        });
      } };
    });
    try {
      const basis = session.confirmedBasis;
      const first = session.submit(command(session, 1, 'editCell', { args: [0, 0, 0, '3'], target: selected }));
      const second = session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '4'], target: selected }, 1));
      expect(session.confirmedBasis).toEqual(basis);
      expect(await first).toMatchObject({ seq: 1, status: 'applied', committed: true });
      expect(await second).toMatchObject({ seq: 2, status: 'applied', committed: true });
      expect(sent.map((item) => item.basis)).toEqual([basis, basis]);
      expect((await session.call.cellInputs(0, 'A1:B1')).cells[0].map((cell) => cell.input)).toEqual(['4', '=A1*2']);
      expect(session.throughSeq).toBe(2);
      expect(session.presentedThroughSeq).toBe(0);
      expect(session.confirmedBasis.version).toBe(await session.call.version());
    } finally { await session.dispose(); }
  });

  test('refusal preserves content history and worker usability', async () => {
    const session = await createTestWorkbookSession(bytes);
    const acks: WorkbookAck[] = [];
    session.onAck((ack) => { acks.push(ack); });
    try {
      await session.submit(command(session, 1, 'editCell', { args: [0, 0, 0, '3'] }));
      const before = await session.save();
      const version = session.confirmedBasis.version;
      const refused = await session.submit(command(session, 2, 'applyEdits', { args: [{
        expectVersion: 'stale', steps: [{ op: 'setCellInputs', target: selected, inputs: [['9']] }],
      }] }));
      expect(refused).toMatchObject({ status: 'refused', reason: 'stale-version', committed: false, affectedSeqs: [2] });
      expect(await session.save()).toEqual(before);
      expect(session.confirmedBasis.version).toBe(version);
      expect(session.failure).toBeUndefined();
      expect(await session.submit(command(session, 3, 'undo', { args: [] }))).toMatchObject({ status: 'applied' });
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('1');
      expect(acks[acks.length - 1]?.editState).toMatchObject({ canUndo: false, canRedo: true });
    } finally { await session.dispose(); }
  });

  test('reports every seq including cancelled dependents', async () => {
    const session = await createTestWorkbookSession(bytes, batchRequests);
    const acks: WorkbookAck[] = [];
    session.onAck((ack) => { acks.push(ack); });
    try {
      const pending = [
        session.submit(command(session, 1, 'editCell', { args: [99, 0, 0, 'bad'] })),
        session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, 'dependent'] }, 1)),
        session.submit(command(session, 3, 'editCell', { args: [0, 0, 0, 'transitive'] }, 2)),
        session.submit(command(session, 4, 'editCell', { args: [0, 0, 0, 'independent'] })),
      ];
      expect((await Promise.all(pending)).map((status) => status.status)).toEqual(['refused', 'cancelled', 'cancelled', 'applied']);
      expect(acks.flatMap((ack) => ack.statuses).map((status) => status.seq)).toEqual([1, 2, 3, 4]);
      expect(session.throughSeq).toBe(4);
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('independent');
    } finally { await session.dispose(); }
  });

  test('deduplicates command delivery', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const input = command(session, 1, 'editCell', { args: [0, 0, 0, '9'] });
      const first = await session.submit(input);
      await session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '10'] }));
      const version = session.confirmedBasis.version;
      expect(await session.submit(input)).toEqual(first);
      expect(session.confirmedBasis.version).toBe(version);
      expect(session.throughSeq).toBe(2);
      await session.submit(command(session, 3, 'undo', { args: [] }));
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('9');
    } finally { await session.dispose(); }
  });

  test('noops and invalid payloads advance throughSeq without moving version', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const before = session.confirmedBasis.version;
      expect(await session.submit(command(session, 1, 'editCell', { args: [0, 0, 0, '1'] })))
        .toMatchObject({ status: 'noop', committed: false });
      expect(await session.submit({ ...command(session, 2, 'undo', { args: [] }),
        payload: { args: ['invalid'] } } as unknown as WorkbookCommand))
        .toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
      expect(session.throughSeq).toBe(2);
      expect(session.confirmedBasis.version).toBe(before);
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('refuses locked targets before changing content or history', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const saved = await session.save();
      expect(await session.submit(command(session, 1, 'editCell', { args: [0, 0, 4, 'locked'] })))
        .toMatchObject({ status: 'refused', reason: 'locked-target', committed: false });
      expect(await session.save()).toEqual(saved);
      expect(await session.submit(command(session, 2, 'undo', { args: [] }))).toMatchObject({ status: 'noop' });
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('rejects wrong identity and epoch without consuming the sequence', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const input = command(session, 1, 'editCell', { args: [0, 0, 0, '3'] });
      await expect(session.call.command({ ...input, incarnation: 'retired' })).rejects.toThrow();
      await expect(session.call.command({ ...input, basis: { ...input.basis, epoch: input.basis.epoch + 1 } }))
        .rejects.toThrow();
      expect(await session.submit(input)).toMatchObject({ status: 'applied' });
      expect(session.throughSeq).toBe(1);
    } finally { await session.dispose(); }
  });

  test('known precommit op validation errors refuse and leave the worker usable', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const saved = await session.save();
      let seq = 0;
      for (const ops of [
        [{ type: 'notAnOperation' }],
        [{ type: 'insertRows', sheet: 0, at: 0, count: 0 }],
        [{ type: 'renameSheet', sheet: 1, name: 'Data' }],
        [{ type: 'removeSheet', index: 1 }, { type: 'removeSheet', index: 0 }],
        [{ type: 'setHyperlinks', sheet: 0, hyperlinks: [{
          range: { start: { row: 1, col: 0 }, end: { row: 1, col: 0 } },
        }] }],
      ]) {
        expect(await session.submit(command(session, ++seq, 'applyOps', { args: [ops] })))
          .toMatchObject({ status: 'refused', committed: false });
      }
      expect(await session.save()).toEqual(saved);
      expect(await session.submit(command(session, ++seq, 'editCell', { args: [0, 0, 0, '2'] })))
        .toMatchObject({ status: 'applied' });
      await expect(session.call.applyEdits({ expectVersion: session.confirmedBasis.version, steps: [] })).rejects.toThrow();
    } finally { await session.dispose(); }
  });

  test('native structural refusal leaves earlier staged operations uncommitted', async () => {
    const calculation = { nowSerial: 46_000.5, randSeed: 123456789 };
    const { session, handle } = await trackedSession(calculation);
    const contexts = spyOn(handle, 'setCalculationContext');
    try {
      await session.submit(command(session, 1, 'applyOps', { args: [[{ type: 'setCell', sheet: 0,
        at: { row: 1, col: 0 }, cell: { value: { kind: 'empty' }, formula: 'SUM(', style: null } }]] }));
      const saved = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      const override = { nowSerial: 50_000.25, randSeed: 987654321 };
      expect(await session.submit(command(session, 2, 'applyOps', { args: [[
        { type: 'renameSheet', sheet: 1, name: 'Uncommitted' },
        { type: 'insertRows', sheet: 0, at: 0, count: 1 },
      ]], calculation: override }))).toMatchObject({ status: 'refused', reason: 'unsupported', committed: false });
      expect(await session.save()).toEqual(saved);
      expect(handle.version()).toBe(version);
      expect(handle.historyState()).toEqual(history);
      expect(contexts.mock.calls).toEqual([[override], [calculation]]);
      expect(session.state.sheets[1].name).toBe('Other');
      expect(await session.submit(command(session, 3, 'editCell', { args: [0, 2, 0, '=NOW()'] })))
        .toMatchObject({ status: 'applied' });
      expect(handle.readCells({ ranges: [{ ...selected, range: { kind: 'a1', a1: 'A3' } }] }))
        .toMatchObject({ ok: true, ranges: [{ cells: [[{ value: { kind: 'number', value: calculation.nowSerial } }]] }] });
      expect(session.failure).toBeUndefined();
    } finally { contexts.mockRestore(); await session.dispose(); }
  });

  for (const kind of ['applyOps', 'moveChart'] as const) {
    test(`native shared drawing ${kind} refusal leaves both sheets and the worker unchanged`, async () => {
      const zip = await JSZip.loadAsync(bytes);
      const columns = kind === 'moveChart' ? '<cols><col min="1" max="6" width="0" customWidth="1"/></cols>' : '';
      zip.file('xl/worksheets/sheet2.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${columns}<sheetData/><drawing r:id="rId1"/></worksheet>`);
      zip.file('xl/worksheets/_rels/sheet2.xml.rels', await zip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string'));
      const session = await createTestWorkbookSession(await zip.generateAsync({ type: 'uint8array' }));
      try {
        const saved = await session.save();
        const version = session.confirmedBasis.version;
        const intent = kind === 'applyOps'
          ? command(session, 1, 'applyOps', { args: [[{ type: 'insertRows', sheet: 0, at: 4, count: 1 }]] })
          : command(session, 1, 'moveChart', { args: [0, 'xl/drawings/drawing1.xml#0', 24, 12] });
        expect(await session.submit(intent))
          .toMatchObject({ status: 'refused', reason: kind === 'applyOps' ? 'unsupported' : 'invalid-step', committed: false });
        expect(await session.save()).toEqual(saved);
        expect(session.confirmedBasis.version).toBe(version);
        expect(await session.call.version()).toBe(version);
        expect(await session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '2'] })))
          .toMatchObject({ status: 'applied', committed: true });
        expect(session.failure).toBeUndefined();
      } finally { await session.dispose(); }
    });
  }

  test('native chart frame drift refuses an anchor recorded before a move', async () => {
    const session = await createTestWorkbookSession(bytes);
    const anchor = { kind: 'twoCell', from: { col: 0, colOff: 0, row: 3, rowOff: 0 },
      to: { col: 5, colOff: 0, row: 15, rowOff: 0 }, edit_as: 'twoCell' };
    try {
      expect(await session.submit(command(session, 1, 'moveChart', { args: [0, 'xl/drawings/drawing1.xml#0', 24, 12] })))
        .toMatchObject({ status: 'applied', committed: true });
      const saved = await session.save();
      const version = session.confirmedBasis.version;
      expect(await session.submit(command(session, 2, 'applyOps', { args: [[{ type: 'setChartAnchor', sheet: 0,
        frame: 'xl/drawings/drawing1.xml#0', part: 'xl/charts/chart1.xml', from: anchor, to: anchor }]] })))
        .toMatchObject({ status: 'refused', reason: 'stale-target', committed: false });
      expect(await session.save()).toEqual(saved);
      expect(session.confirmedBasis.version).toBe(version);
      expect(await session.call.version()).toBe(version);
      expect(await session.submit(command(session, 3, 'editCell', { args: [0, 0, 0, '2'] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('native forced proposal acceptance refuses text made too long by a format change', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      const proposed = await session.submit(command(session, 1, 'propose', {
        args: ['agent', null, [{ sheet: 0, row: 0, col: 0, input: `="${'x'.repeat(32_765)}"` }]],
      }));
      expect(proposed).toMatchObject({ status: 'applied', committed: false });
      expect(await session.submit(command(session, 2, 'setNumberFormat', { args: [0, 'A1', 'plainText'] })))
        .toMatchObject({ status: 'applied', committed: true });
      const saved = await session.save();
      const version = session.confirmedBasis.version;
      expect(await session.submit(command(session, 3, 'acceptProposal', { args: [(proposed.result as wasm.Proposal).id, { force: true }] })))
        .toMatchObject({ status: 'refused', reason: 'invalid-step', committed: false });
      expect(await session.save()).toEqual(saved);
      expect(session.confirmedBasis.version).toBe(version);
      expect(await session.call.version()).toBe(version);
      expect(await session.submit(command(session, 4, 'editCell', { args: [0, 0, 0, '2'] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  test('native anchor validation refuses unsupported changes and invalid geometry', async () => {
    const session = await createTestWorkbookSession(bytes);
    const from = { kind: 'twoCell', from: { col: 0, colOff: 0, row: 3, rowOff: 0 },
      to: { col: 5, colOff: 0, row: 15, rowOff: 0 }, edit_as: 'twoCell' };
    try {
      let seq = 0;
      for (const [to, reason] of [
        [{ ...from, edit_as: 'oneCell' }, 'unsupported'],
        [{ kind: 'oneCell', from: from.from, extent: { cx: 800_000, cy: 1_000_000 } }, 'unsupported'],
        [{ ...from, from: { ...from.from, colOff: -1 } }, 'invalid-step'],
        [{ ...from, to: from.from }, 'invalid-step'],
        [{ ...from, to: { ...from.to, col: 16_384 } }, 'invalid-step'],
        [{ ...from, to: { ...from.to, colOff: 2 ** 40 } }, 'invalid-step'],
      ] as const) {
        const saved = await session.save();
        const version = session.confirmedBasis.version;
        expect(await session.submit(command(session, ++seq, 'applyOps', { args: [[{ type: 'setChartAnchor', sheet: 0,
          frame: 'xl/drawings/drawing1.xml#0', part: 'xl/charts/chart1.xml', from, to }]] })))
          .toMatchObject({ status: 'refused', reason, committed: false });
        expect(await session.save()).toEqual(saved);
        expect(session.confirmedBasis.version).toBe(version);
        expect(await session.call.version()).toBe(version);
        expect(await session.submit(command(session, ++seq, 'editCell', { args: [0, 0, 0, String(seq)] })))
          .toMatchObject({ status: 'applied', committed: true });
        expect(session.failure).toBeUndefined();
      }
    } finally { await session.dispose(); }
  });

  for (const field of ['defined name', 'chart reference', 'chart xml'] as const) {
    test(`native rename refuses an invalid ${field} before commit`, async () => {
      const zip = await JSZip.loadAsync(bytes);
      const formula = `Data!$A$1&amp;"${'x'.repeat(32_755)}"`;
      if (field === 'defined name') {
        const workbook = await zip.file('xl/workbook.xml')!.async('string');
        zip.file('xl/workbook.xml', workbook.replace('</workbook>',
          `<definedNames><definedName name="NearLimit">${formula}</definedName></definedNames></workbook>`));
      } else if (field === 'chart reference') {
        const chart = await zip.file('xl/charts/chart1.xml')!.async('string');
        const reference = `(${Array(2_047).fill('Data!$A$1').join(',')},Data!$A1)`;
        zip.file('xl/charts/chart1.xml', chart.replace('Data!$A$1:$B$1', reference)
          .replace(/<c:numCache>.*?<\/c:numCache>/s, ''));
      }
      const session = await createTestWorkbookSession(await zip.generateAsync({ type: 'uint8array' }));
      try {
        const saved = await session.save();
        const version = session.confirmedBasis.version;
        expect(await session.submit(command(session, 1, 'applyOps', {
          args: [[{ type: 'renameSheet', sheet: 0, name: field === 'chart xml' ? 'Data\u0001' : 'LongerData' }]],
        }))).toMatchObject({ status: 'refused', reason: 'invalid-step', committed: false });
        expect(await session.save()).toEqual(saved);
        expect(session.confirmedBasis.version).toBe(version);
        expect(await session.call.version()).toBe(version);
        expect(await session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '2'] })))
          .toMatchObject({ status: 'applied', committed: true });
        expect(session.failure).toBeUndefined();
      } finally { await session.dispose(); }
    });
  }

  test('staged style errors refuse without committing preceding operations', async () => {
    const { session, handle } = await trackedSession();
    try {
      const saved = await session.save();
      const version = handle.version();
      const history = handle.historyState();
      let seq = 0;
      for (const patch of [{ fontSize: 0 }, { fontSize: 401 }, { fontFamily: '' }, { textColor: 'invalid' }, { fillColor: 'invalid' }]) {
        expect(await session.submit(command(session, ++seq, 'applyOps', { args: [[
          { type: 'renameSheet', sheet: 1, name: 'Uncommitted' },
          { type: 'patchRangeStyle', sheet: 0, range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }, patch },
        ]] }))).toMatchObject({ status: 'refused', reason: 'invalid-step', committed: false });
        expect(await session.save()).toEqual(saved);
        expect(handle.version()).toBe(version);
        expect(handle.historyState()).toEqual(history);
      }
      expect(await session.submit(command(session, ++seq, 'patchRangeStyle', { args: [0, 'A1', { fontSize: 0 }] })))
        .toMatchObject({ status: 'refused', reason: 'invalid-step', committed: false });
      expect(handle.historyState()).toEqual(history);
      expect(await session.submit(command(session, ++seq, 'editCell', { args: [0, 0, 0, '2'] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect(session.failure).toBeUndefined();
    } finally { await session.dispose(); }
  });

  for (const refusal of [
    { method: 'setNumberFormat', message: 'number format table is full', reason: 'limit-exceeded' },
    { method: 'applyOps', message: 'formula at sheet 0, B1 cannot be safely rewritten', reason: 'unsupported' },
    { method: 'applyOps', message: 'defined name Dynamic cannot be safely rewritten', reason: 'unsupported' },
    { method: 'applyOps', message: 'chart reference in xl/charts/chart1.xml cannot be safely rewritten', reason: 'unsupported' },
    { method: 'applyOps', message: 'xl/tables/table1.xml references cells this edit would move, and it cannot be rewritten', reason: 'unsupported' },
  ] as const) {
    for (const fault of ['precommit', 'postcommit', 'trap'] as const) {
      test(`${fault} ${refusal.message} retains refusal and context guards`, async () => {
        const calculation = { nowSerial: 46_000.5, randSeed: 123456789 };
        const override = { nowSerial: 50_000.25, randSeed: 987654321 };
        const { session, handle } = await trackedSession(calculation);
        const contexts = spyOn(handle, 'setCalculationContext');
        const failing = spyOn(handle, refusal.method).mockImplementation(() => {
          if (fault === 'postcommit') handle.editCell(0, 2, 0, 'committed');
          throw fault === 'trap' ? new WebAssembly.RuntimeError(refusal.message) : new Error(refusal.message);
        });
        try {
          const saved = await session.save();
          const version = handle.version();
          const history = handle.historyState();
          const intent = refusal.method === 'setNumberFormat'
            ? command(session, 1, 'setNumberFormat', { args: [0, 'A1', { type: 'custom', pattern: '0.0000' }], calculation: override })
            : command(session, 1, 'applyOps', { args: [[{ type: 'insertRows', sheet: 0, at: 0, count: 1 }]], calculation: override });
          if (fault === 'precommit') {
            expect(await session.submit(intent)).toMatchObject({ status: 'refused', reason: refusal.reason, committed: false });
            expect(contexts.mock.calls).toEqual([[override], [calculation]]);
            expect(await session.save()).toEqual(saved);
            expect(handle.version()).toBe(version);
            expect(handle.historyState()).toEqual(history);
            failing.mockRestore();
            expect(await session.submit(command(session, 2, 'editCell', { args: [0, 2, 0, '=NOW()'] })))
              .toMatchObject({ status: 'applied' });
            expect(handle.readCells({ ranges: [{ ...selected, range: { kind: 'a1', a1: 'A3' } }] }))
              .toMatchObject({ ok: true, ranges: [{ cells: [[{ value: { kind: 'number', value: calculation.nowSerial } }]] }] });
            expect(session.failure).toBeUndefined();
          } else {
            await expect(session.submit(intent)).rejects.toBeInstanceOf(SessionFailure);
            expect(session.failure?.code).toBe(fault === 'trap' ? 'trap' : 'crash');
            expect(contexts.mock.calls).toEqual([[override]]);
          }
        } finally { failing.mockRestore(); contexts.mockRestore(); await session.dispose(); }
      });
    }
  }

  for (const fault of [
    { error: new WebAssembly.RuntimeError('bad ops: trap'), committed: false, code: 'trap' },
    { error: new Error('bad ops: failure after commit'), committed: true, code: 'crash' },
  ] as const) {
    test(`does not turn a ${fault.code} into a safe op refusal`, async () => {
      const { session, handle } = await trackedSession();
      const apply = handle.applyOps.bind(handle);
      const failing = spyOn(handle, 'applyOps').mockImplementation((...args) => {
        if (fault.committed) apply(...args);
        throw fault.error;
      });
      try {
        await expect(session.submit(command(session, 1, 'applyOps', {
          args: [[{ type: 'renameSheet', sheet: 1, name: 'Renamed' }]],
        }))).rejects.toBeInstanceOf(SessionFailure);
        expect(session.failure?.code).toBe(fault.code);
        await expect(session.save()).rejects.toBeInstanceOf(SessionFailure);
      } finally { failing.mockRestore(); await session.dispose(); }
    });
  }

  test('applies worker-captured formatting without a peer', async () => {
    const session = await createTestWorkbookSession(bytes);
    try {
      await session.submit(command(session, 1, 'patchRangeStyle', { args: [0, 'B1', { bold: true }] }));
      const captured = await session.call.captureFormat(0, 'B1');
      expect(captured).toMatchObject({ rows: 1, columns: 1 });
      expect(await session.submit(command(session, 2, 'applyFormat', {
        target: selected, capturedFrom: { ...selected, range: { kind: 'a1', a1: 'B1' } },
      }))).toMatchObject({ status: 'applied', committed: true });
      await session.submit(command(session, 3, 'select', { selection: selected }));
      expect(session.editState.formatting?.bold).toBe(true);
      await session.submit(command(session, 4, 'patchRangeStyle', { args: [0, 'B1', { bold: false }] }));
      await session.submit(command(session, 5, 'applyFormat', { args: [0, 'A2', captured] }));
      await session.submit(command(session, 6, 'select', { selection: {
        ...selected, range: { kind: 'a1', a1: 'A2' },
      } }));
      expect(session.editState.formatting?.bold).toBe(true);
    } finally { await session.dispose(); }
  });

  test('formats merged ranges through every formatting mutator', async () => {
    const { session, handle } = await trackedSession();
    const main = wasm.openWorkbook(bytes);
    const validating = spyOn(handle, 'validateEdits');
    try {
      main.patchRangeStyle(0, 'B1', { bold: true });
      const format = main.captureFormat(0, 'B1');
      let seq = 0;
      for (const kind of ['patchRangeStyle', 'setNumberFormat', 'applyFormat'] as const) {
        const payload: WorkbookCommandPayload<typeof kind> = kind === 'patchRangeStyle'
          ? { args: [0, 'D1:E1', { bold: true }] } : kind === 'setNumberFormat'
          ? { args: [0, 'D1:E1', 'percent'] } : { args: [0, 'D1:E1', format] };
        const status = await session.submit(command(session, ++seq, kind, payload));
        expect(status).toMatchObject({ status: 'applied', committed: true });
        expect(handle.selectionFormatting(0, 'D1:E1')).toMatchObject(kind === 'setNumberFormat'
          ? { numberFormat: 'percent' } : { bold: true });
        expect(handle.mergedRanges(0, 'D1:E1')).toEqual([{ start: { row: 0, col: 3 }, end: { row: 0, col: 4 } }]);
        await session.submit(command(session, ++seq, 'undo', { args: [] }));
        expect(handle.selectionFormatting(0, 'D1:E1')).toMatchObject({ bold: false, numberFormat: 'automatic' });
      }
      expect(validating).not.toHaveBeenCalled();
      expect(session.failure).toBeUndefined();
    } finally { validating.mockRestore(); main.dispose(); await session.dispose(); }
  });

  test('save fences accepted commands and ordered reads', async () => {
    const session = await createTestWorkbookSession(bytes, batchRequests);
    try {
      const first = session.submit(command(session, 1, 'editCell', { args: [0, 0, 0, '11'] }));
      const second = session.submit(command(session, 2, 'editCell', { args: [0, 0, 1, '=A1+1'] }, 1));
      const saved = session.save();
      const read = session.call.cellInputs(0, 'A1:B1');
      const reopened = wasm.openWorkbook(await saved);
      try {
        expect(reopened.cell(0, 0, 0).input).toBe('11');
        expect(reopened.cell(0, 0, 1).input).toBe('=A1+1');
        expect((await read).cells[0][0].input).toBe('11');
        await Promise.all([first, second]);
      } finally { reopened.dispose(); }
    } finally { await session.dispose(); }
  });

  test('pushes formatting history calculation and catalog changes without unchanged fields', async () => {
    const session = await createTestWorkbookSession(bytes);
    const acks: WorkbookAck[] = [];
    session.onAck((ack) => { acks.push(ack); });
    try {
      await session.submit(command(session, 1, 'select', { selection: selected }));
      expect(acks[acks.length - 1]?.editState).toMatchObject({ canUndo: false, canRedo: false, formatting: { bold: false } });
      expect(acks[acks.length - 1]?.editState?.calculation).toBeDefined();
      expect(acks[acks.length - 1]?.editState?.sheets).toHaveLength(2);
      await session.submit(command(session, 2, 'patchRangeStyle', { args: [0, 'A1', { bold: true }] }));
      expect(acks[acks.length - 1]?.editState).toMatchObject({ canUndo: true, formatting: { bold: true } });
      expect(acks[acks.length - 1]?.editState?.sheets).toBeUndefined();
      await session.submit(command(session, 3, 'select', { selection: selected }));
      expect(acks[acks.length - 1]?.editState).toBeUndefined();
      await session.submit(command(session, 4, 'applyOps', { args: [[{ type: 'renameSheet', sheet: 0, name: 'Renamed' }]] }));
      expect(acks[acks.length - 1]?.editState?.sheets?.[0].name).toBe('Renamed');
    } finally { await session.dispose(); }
  });

  test('content-only Enter commits reuse selection formatting merges and catalogs', async () => {
    const { session, handle } = await trackedSession();
    await session.submit(command(session, 1, 'select', { selection: selected }));
    const formatting = spyOn(handle, 'selectionFormatting');
    const merges = spyOn(handle, 'mergedRanges');
    const catalog = spyOn(handle, 'readCells');
    const info = spyOn(handle, 'sheetInfo');
    const acks: WorkbookAck[] = [];
    session.onAck((ack) => { acks.push(ack); });
    try {
      for (let seq = 2; seq <= 21; seq++) {
        expect(await session.submit(command(session, seq, 'editCell', { args: [0, 1, 0, String(seq)] })))
          .toMatchObject({ status: 'applied', committed: true });
      }
      expect(formatting).not.toHaveBeenCalled();
      expect(merges).not.toHaveBeenCalled();
      expect(catalog).not.toHaveBeenCalled();
      expect(info).not.toHaveBeenCalled();
      expect(acks.every((ack) => !ack.editState?.formatting && !ack.editState?.mergedRanges && !ack.editState?.sheets)).toBe(true);
      await session.submit(command(session, 22, 'patchRangeStyle', { args: [0, 'A1', { bold: true }] }));
      expect(formatting).toHaveBeenCalledTimes(1);
      expect(merges).not.toHaveBeenCalled();
      expect(session.editState.formatting?.bold).toBe(true);
      await session.submit(command(session, 23, 'applyOps', { args: [[
        { type: 'mergeCells', sheet: 0, range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } } },
      ]] }));
      expect(merges).toHaveBeenCalledTimes(1);
      expect(catalog).toHaveBeenCalledTimes(1);
      expect(session.editState.mergedRanges).toEqual([{ start: { row: 0, col: 0 }, end: { row: 0, col: 1 } }]);
      await session.submit(command(session, 24, 'undo', { args: [] }));
      expect(session.editState.mergedRanges).toEqual([]);
    } finally { info.mockRestore(); catalog.mockRestore(); merges.mockRestore(); formatting.mockRestore(); await session.dispose(); }
  });

  test('history anchors append without scans and truncate only discarded redo entries', async () => {
    const handle = wasm.openWorkbook(bytes);
    const identity = { session: 'history', incarnation: 'worker' };
    const state = createWorkbookCommandState(identity, { epoch: 1, version: handle.version() }, selected, handle.sheetInfo().sheetIds);
    const truncations: number[] = [];
    state.historySelections = new Proxy(state.historySelections, {
      ownKeys() { throw new Error('History anchors must not be scanned'); },
      get(target, property, receiver) {
        if (property === Symbol.iterator || ['keys', 'values', 'entries', 'forEach'].includes(String(property))) {
          throw new Error('History anchors must not be traversed');
        }
        return Reflect.get(target, property, receiver);
      },
      set(target, property, value, receiver) {
        if (property === 'length') truncations.push(value);
        return Reflect.set(target, property, value, receiver);
      },
    });
    let seq = 0;
    const host = createWorkbookCommandHost({ handle, state,
      frame: () => { throw new Error('Unexpected history frame'); } });
    const send = async <K extends WorkbookCommand['kind']>(kind: K, payload: WorkbookCommandPayload<K>) => {
      const result = await host.command({ ...identity, seq: ++seq, basis: { epoch: 1, version: handle.version() },
        kind, payload } as WorkbookCommand);
      return isTransferResult(result) ? result.value : result;
    };
    try {
      for (let at = 0; at < 40; at++) {
        await send('editCell', { args: [0, 1, 0, String(at)], selection: { ...selected, range: { kind: 'a1', a1: `A${at + 2}` } } });
      }
      expect(truncations).toEqual([]);
      await send('undo', { args: [] });
      await send('undo', { args: [] });
      const branchSelection = structuredClone(state.selection);
      const after = { ...selected, range: { kind: 'a1' as const, a1: 'C3' } };
      await send('editCell', { args: [0, 1, 0, 'branch'], selection: after });
      expect(truncations).toEqual([39]);
      expect((await send('redo', { args: [] })).statuses[0].status).toBe('noop');
      await send('undo', { args: [] });
      expect(state.selection).toEqual(branchSelection);
      await send('redo', { args: [] });
      expect(state.selection).toEqual(after);
      expect(handle.cell(0, 1, 0).input).toBe('branch');
    } finally { handle.dispose(); }
  });

  test('undo redo restore worker selection and history availability', async () => {
    const session = await createTestWorkbookSession(bytes);
    const acks: WorkbookAck[] = [];
    session.onAck((ack) => { acks.push(ack); });
    try {
      await session.submit(command(session, 1, 'select', { selection: selected }));
      await session.submit(command(session, 2, 'editCell', { args: [0, 0, 0, '5'], selection: {
        ...selected, range: { kind: 'a1', a1: 'A2' },
      } }));
      await session.submit(command(session, 3, 'undo', { args: [] }));
      expect(acks[acks.length - 1]?.selection).toEqual(selected);
      expect(acks[acks.length - 1]?.editState).toMatchObject({ canUndo: false, canRedo: true });
      await session.submit(command(session, 4, 'redo', { args: [] }));
      expect(acks[acks.length - 1]?.selection?.range).toEqual({ kind: 'a1', a1: 'A2' });
      expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe('5');
    } finally { await session.dispose(); }
  });

  test('setReadOnly is an ordered mode fence', async () => {
    const session = await createTestWorkbookSession(bytes, batchRequests);
    try {
      const commands = [
        command(session, 1, 'editCell', { args: [0, 0, 0, '2'] }),
        command(session, 2, 'setReadOnly', { readOnly: true }),
        command(session, 3, 'undo', { args: [] }),
        command(session, 4, 'select', { selection: { ...selected, range: { kind: 'a1', a1: 'A2' } } }),
        command(session, 5, 'setReadOnly', { readOnly: false }),
        command(session, 6, 'undo', { args: [] }),
      ];
      expect((await Promise.all(commands.map((item) => session.submit(item)))).map((status) => status.status))
        .toEqual(['applied', 'applied', 'refused', 'applied', 'applied', 'applied']);
    } finally { await session.dispose(); }
  });

  test('unknown failure after mutation starts is terminal', async () => {
    const original = wasm.openWorkbook;
    let handle: wasm.WorkbookHandle | undefined;
    const opening = spyOn(wasm, 'openWorkbook').mockImplementation((...args) => {
      handle = original(...args);
      return handle;
    });
    let session: WorkbookSession;
    try { session = await createTestWorkbookSession(bytes); }
    finally { opening.mockRestore(); }
    if (!handle) throw new Error('Missing worker handle');
    const edit = handle.editCell.bind(handle);
    const failing = spyOn(handle, 'editCell').mockImplementation((...args) => {
      edit(...args);
      throw new Error('post-commit failure');
    });
    try {
      await expect(session.submit(command(session, 1, 'editCell', { args: [0, 0, 0, '8'] })))
        .rejects.toBeInstanceOf(SessionFailure);
      expect(session.failure?.code).toBe('crash');
      await expect(session.save()).rejects.toBeInstanceOf(SessionFailure);
    } finally { failing.mockRestore(); await session.dispose(); }
  });

  test('copy origin errors from batch edits are classified as invalid commands', () => {
    expect(classifyRefusal(new Error('bad edit args: copy origin is outside the grid'), 'editCells'))
      .toMatchObject({ reason: 'invalid-command' });
  });

  test('out-of-grid copy origins are classified refusals without changing state', async () => {
    const session = await createTestWorkbookSession(plainBytes);
    try {
      const before = await session.save();
      for (const [index, copyFrom] of [
        { row: 1_048_576, col: 0 }, { row: 0, col: 16_384 },
        { row: Number.MAX_SAFE_INTEGER, col: 0 },
      ].entries()) {
        expect(await session.submit(command(session, index + 1, 'editCells', {
          args: [0, [{ row: 2, col: 2, input: '=A1', copyFrom }]],
        }))).toMatchObject({ status: 'refused', reason: 'invalid-command', committed: false });
        expect(await session.save()).toEqual(before);
        expect(session.failure).toBeUndefined();
      }
      expect(await session.submit(command(session, 4, 'editCell', { args: [0, 2, 2, 'valid'] })))
        .toMatchObject({ status: 'applied', committed: true });
    } finally { await session.dispose(); }
  });

  test('copied formula origins survive worker command validation and undo', async () => {
    const session = await createTestWorkbookSession(plainBytes);
    try {
      expect(await session.submit(command(session, 1, 'editCells', {
        args: [0, [
          { row: 2, col: 2, input: '=A1+$B$1+$A1+A$1', copyFrom: { row: 0, col: 2 } },
          { row: 3, col: 2, input: '=SUM(A1:B2)', copyFrom: { row: 0, col: 2 } },
          { row: 2, col: 3, input: '=A1' },
        ]],
      }))).toMatchObject({ status: 'applied', committed: true });
      const inputs = await session.call.cellInputs(0, 'C3:D4');
      expect(inputs.cells.map((row) => row.map((cell) => cell.input))).toEqual([
        ['=A3+$B$1+$A3+A$1', '=A1'], ['=SUM(A4:B5)', ''],
      ]);
      expect(await session.submit(command(session, 2, 'undo', { args: [] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect((await session.call.cellInputs(0, 'C3')).cells[0][0].input).toBe('');
    } finally { await session.dispose(); }
  });

  test('moveRange rewrites dependent formulas and undoes through worker commands', async () => {
    const session = await createTestWorkbookSession(plainBytes);
    try {
      expect(await session.submit(command(session, 1, 'moveRange', {
        args: [0, 'A1', 'C3'],
      }))).toMatchObject({ status: 'applied', committed: true });
      expect((await session.call.cellInputs(0, 'A1:C3')).cells).toEqual([
        [{ a1: 'A1', input: '', isFormula: false }, { a1: 'B1', input: '=C3*2', isFormula: true }, { a1: 'C1', input: '', isFormula: false }],
        [{ a1: 'A2', input: '', isFormula: false }, { a1: 'B2', input: '', isFormula: false }, { a1: 'C2', input: '', isFormula: false }],
        [{ a1: 'A3', input: '', isFormula: false }, { a1: 'B3', input: '', isFormula: false }, { a1: 'C3', input: '1', isFormula: false }],
      ]);
      expect(await session.submit(command(session, 2, 'editCell', {
        args: [0, 0, 2, '=SUM(B1:C3)'],
      }))).toMatchObject({ status: 'applied', committed: true });
      expect(await session.submit(command(session, 3, 'moveRange', {
        args: [0, 'C3', 'D4'],
      }))).toMatchObject({ status: 'refused', reason: 'unsupported', committed: false });
      expect(await session.submit(command(session, 4, 'undo', { args: [] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect(await session.submit(command(session, 5, 'undo', { args: [] })))
        .toMatchObject({ status: 'applied', committed: true });
      expect((await session.call.cellInputs(0, 'A1:B1')).cells[0]).toEqual([
        { a1: 'A1', input: '1', isFormula: false },
        { a1: 'B1', input: '=A1*2', isFormula: true },
      ]);
      expect((await session.call.cellInputs(0, 'C3')).cells[0][0].input).toBe('');
    } finally { await session.dispose(); }
  });

  for (const kind of Object.keys(WORKBOOK_REPLAY_MUTATORS) as WorkbookReplayMethod[]) {
    test(`applies replay mutator ${kind} without an edit peer`, async () => {
      const source = kind === 'moveRange' ? plainBytes : bytes;
      const session = await createTestWorkbookSession(source);
      let seq = 0;
      const send = (op: WorkbookCommand) => session.submit(op);
      const main = wasm.openWorkbook(source);
      try {
        let payload: WorkbookCommandPayload<WorkbookReplayMethod>;
        switch (kind) {
          case 'editCell': payload = { args: [0, 0, 0, '7'] }; break;
          case 'editCells': payload = { args: [0, [{ row: 0, col: 0, input: '7' }]] }; break;
          case 'applyEdits': payload = { args: [{ expectVersion: session.confirmedBasis.version,
            steps: [{ op: 'setCellInputs', target: selected, inputs: [['7']] }] }] }; break;
          case 'applyOps': payload = { args: [[{ type: 'insertRows', sheet: 0, at: 3, count: 1 }]] }; break;
          case 'patchRangeStyle': payload = { args: [0, 'A1', { bold: true }] }; break;
          case 'setNumberFormat': payload = { args: [0, 'A1', 'percent'] }; break;
          case 'applyFormat':
            main.patchRangeStyle(0, 'B1', { bold: true });
            payload = { args: [0, 'A1', main.captureFormat(0, 'B1')] };
            break;
          case 'moveChart': {
            const chart = main.displayList({ x: 0, y: 0, width: 800, height: 800 }).charts?.find((item) => item.movable);
            if (!chart) throw new Error('Missing synthetic movable chart');
            payload = { args: [0, chart.id, 24, 12] };
            break;
          }
          case 'moveRange': payload = { args: [0, 'A1', 'A2'] }; break;
          case 'setActiveSheet': payload = { args: [1] }; break;
          case 'undo':
          case 'redo':
            await send(command(session, ++seq, 'editCell', { args: [0, 0, 0, '7'] }));
            if (kind === 'redo') await send(command(session, ++seq, 'undo', { args: [] }));
            payload = { args: [] };
            break;
          case 'propose': payload = { args: ['agent', null, [{ sheet: 0, row: 0, col: 0, input: '7' }]] }; break;
          case 'acceptProposal':
          case 'rejectProposal': {
            const proposed = await send(command(session, ++seq, 'propose', {
              args: ['agent', null, [{ sheet: 0, row: 0, col: 0, input: '7' }]],
            }));
            payload = { args: [(proposed.result as wasm.Proposal).id] };
            break;
          }
        }
        const before = session.confirmedBasis.version;
        const status = await send(command(session, ++seq, kind, payload));
        expect(status.status).toBe('applied');
        expect(session.failure).toBeUndefined();
        expect(session.throughSeq).toBe(seq);
        if (kind === 'setActiveSheet' || kind === 'propose' || kind === 'rejectProposal') {
          expect(status.committed).toBe(false);
          expect(session.confirmedBasis.version).toBe(before);
        } else {
          expect(status.committed).toBe(true);
          expect(session.confirmedBasis.version).not.toBe(before);
        }
        if (['editCell', 'editCells', 'applyEdits', 'acceptProposal', 'redo', 'undo'].includes(kind)) {
          expect((await session.call.cellInputs(0, 'A1')).cells[0][0].input).toBe(kind === 'undo' ? '1' : '7');
        }
        if (kind === 'moveRange') expect((await session.call.cellInputs(0, 'A2')).cells[0][0].input).toBe('1');
        if (kind === 'patchRangeStyle' || kind === 'applyFormat') expect(session.editState.formatting?.bold).toBe(true);
        if (kind === 'setNumberFormat') expect(session.editState.formatting?.numberFormat).toBe('percent');
        if (kind === 'setActiveSheet') expect(session.state.activeSheet).toBe(1);
        if (kind === 'rejectProposal') expect(status.result).toBe(true);
      } finally { main.dispose(); await session.dispose(); }
    });
  }
});
