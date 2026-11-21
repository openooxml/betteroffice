import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { engineKind } from '../wasm/engine';

async function workbook(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Protected" sheetId="1" r:id="rId1"/><sheet name="Unprotected" sheetId="2" r:id="rId2"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('xl/styles.xml', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf applyProtection="1"><protection locked="0"/></xf></cellXfs></styleSheet>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row></sheetData><sheetProtection sheet="1"/></worksheet>');
  zip.file('xl/worksheets/sheet2.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData></worksheet>');
  return zip.generateAsync({ type: 'uint8array' });
}

test(`protected cell input on the ${engineKind} engine`, async () => {
  const { initWasm } = await import('../wasm/loader');
  const { createTestWorkbookSession } = await import('./testHelpers');
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const session = await createTestWorkbookSession(await workbook(), undefined, {
    includeEditState: true, calculation: { nowSerial: 45000.75, randSeed: 42 },
  });
  let seq = 0;
  const edit = (sheet: number, col: number, input: string) => session.submit({
    ...session.commandIdentity, seq: ++seq, basis: session.confirmedBasis,
    kind: 'editCell', payload: { args: [sheet, 0, col, input] },
  });
  const inputs = async () => [
    ...(await session.call.cellInputs(0, 'A1:B1')).cells[0].map((cell) => cell.input),
    (await session.call.cellInputs(1, 'A1')).cells[0][0].input,
  ];
  try {
    const before = session.confirmedBasis.version;
    expect(await inputs()).toEqual(['1', '2', '3']);
    const locked = await edit(0, 0, 'locked');
    expect(locked).toMatchObject({ status: 'refused', reason: 'locked-target', committed: false,
      result: { ok: false, version: before, failure: { code: 'locked-target',
        message: 'A1 is locked on protected sheet "Protected"' } } });
    expect(await inputs()).toEqual(['1', '2', '3']);
    expect(session.confirmedBasis.version).toBe(before);
    const unlocked = await edit(0, 1, 'unlocked');
    expect(unlocked).toMatchObject({ status: 'applied', committed: true });
    expect(await inputs()).toEqual(['1', 'unlocked', '3']);
    expect(session.confirmedBasis.version).not.toBe(before);
    const afterUnlocked = session.confirmedBasis.version;
    const unprotected = await edit(1, 0, 'unprotected');
    expect(unprotected).toMatchObject({ status: 'applied', committed: true });
    expect(await inputs()).toEqual(['1', 'unlocked', 'unprotected']);
    expect(session.confirmedBasis.version).not.toBe(afterUnlocked);
    expect(session.failure).toBeUndefined();
  } finally { await session.dispose(); }
}, 30_000);

test('protected cell input refreshes an editable sheet cache before catalog projection', async () => {
  const { initWasm, openWorkbook } = await import('../wasm/loader');
  const { applyWorkbookCommandIntent, createWorkbookCommandState, projectEditState } = await import('./commandHandlers');
  await initWasm(await readFile(new URL('../wasm/generated/xlsx_wasm_bg.wasm', import.meta.url)));
  const handle = openWorkbook(await workbook());
  const identity = { session: 'protected-catalog', incarnation: 'worker' };
  const basis = { epoch: 1, version: handle.version() };
  const sheetIds = handle.sheetInfo().sheetIds;
  const state = createWorkbookCommandState(identity, basis,
    { sheetId: sheetIds[1], range: { kind: 'a1', a1: 'A1' } }, sheetIds);
  let pending = state;
  const sheetInfo = handle.sheetInfo.bind(handle);
  const catalog = spyOn(handle, 'sheetInfo').mockImplementation(() => {
    if (state.catalogDirty && pending === state) pending = structuredClone(state);
    return sheetInfo();
  });
  try {
    projectEditState(handle, state);
    expect(state.editableSheets).toEqual([false, true]);
    const moved = applyWorkbookCommandIntent(handle, state, { ...identity, seq: 1, basis,
      kind: 'applyOps', payload: { args: [[{ type: 'moveSheet', from: 1, to: 0 }]] } });
    expect(moved.statuses[0]).toMatchObject({ status: 'applied', committed: true });
    expect(pending).not.toBe(state);
    expect(pending.catalogDirty).toBe(true);
    expect(pending.editableSheets).toEqual([false, true]);
    expect(state.editableSheets).toEqual([true, false]);
    expect(handle.sheetInfo().sheetNames).toEqual(['Unprotected', 'Protected']);
    const before = handle.version();
    const locked = applyWorkbookCommandIntent(handle, pending, { ...identity, seq: 2, basis,
      kind: 'editCell', payload: { args: [1, 0, 0, 'locked'],
        target: { sheetId: sheetIds[0], range: { kind: 'a1', a1: 'A1' } } } });
    expect(locked.statuses[0]).toMatchObject({ status: 'refused', reason: 'locked-target', committed: false,
      result: { ok: false, version: before, failure: { code: 'locked-target',
        message: 'A1 is locked on protected sheet "Protected"' } } });
    expect(handle.cell(1, 0, 0).input).toBe('1');
    expect(handle.version()).toBe(before);
  } finally { catalog.mockRestore(); handle.dispose(); }
});
