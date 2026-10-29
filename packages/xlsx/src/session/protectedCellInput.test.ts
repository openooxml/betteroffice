import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import type { XlsxEditRefusal } from '../edits';
import { packageVersion, selectBackend, type EngineBackend } from '../wasm/engine';

const CHILD = 'XLSX_PROTECTED_CELL_INPUT_TEST';
const RESULT = 'XLSX_PROTECTED_CELL_INPUT_OUTCOME=';
const nativeUnavailable = (() => {
  try {
    selectBackend({ env: 'native', packageVersion, runtime: {
      hasProcess: true, hasGetBuiltinModule: typeof process.getBuiltinModule === 'function',
      hasNodeVersion: typeof process.versions.node === 'string' && process.versions.node.length > 0,
      hasDom: false,
    }, loadNative: () => createRequire(new URL('../wasm/engine.ts', import.meta.url))('@betteroffice/xlsx-native') });
    return undefined;
  } catch (error) { return error instanceof Error ? error.message : String(error); }
})();

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

async function protectedCellOutcomes(engine: EngineBackend) {
  const { engineBackend, initWasm } = await import('../wasm/loader');
  const { createTestWorkbookSession } = await import('./testHelpers');
  expect(engineBackend()).toBe(engine);
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
    const { version: _, ...refusal } = locked.result as XlsxEditRefusal;
    const unlocked = await edit(0, 1, 'unlocked');
    expect(unlocked).toMatchObject({ status: 'applied', committed: true });
    expect(await inputs()).toEqual(['1', 'unlocked', '3']);
    expect(session.confirmedBasis.version).not.toBe(before);
    const afterUnlocked = session.confirmedBasis.version;
    const unprotected = await edit(1, 0, 'unprotected');
    expect(unprotected).toMatchObject({ status: 'applied', committed: true });
    const values = await inputs();
    expect(values).toEqual(['1', 'unlocked', 'unprotected']);
    expect(session.confirmedBasis.version).not.toBe(afterUnlocked);
    expect(session.failure).toBeUndefined();
    return { locked: { ...locked, result: refusal }, unlocked, unprotected, values };
  } finally { await session.dispose(); }
}

type ProtectionOutcomes = Awaited<ReturnType<typeof protectedCellOutcomes>>;
const outcomes = new Map<EngineBackend, Promise<ProtectionOutcomes>>();

function engineOutcomes(engine: EngineBackend): Promise<ProtectionOutcomes> {
  let result = outcomes.get(engine);
  if (!result) {
    result = (async () => {
      const child = Bun.spawn([process.execPath, 'test', import.meta.path, '-t', `^protected cell input on ${engine}$`], {
        env: { ...process.env, BETTEROFFICE_ENGINE: engine, [CHILD]: engine }, stdout: 'pipe', stderr: 'pipe',
      });
      const deadline = setTimeout(() => child.kill(), 20_000);
      try {
        const [status, stdout, stderr] = await Promise.all([
          child.exited, new Response(child.stdout).text(), new Response(child.stderr).text(),
        ]);
        expect({ status, output: status === 0 ? '' : stdout + stderr }).toEqual({ status: 0, output: '' });
        const line = stdout.split('\n').find((line) => line.startsWith(RESULT));
        if (!line) throw new Error(`Missing ${engine} protection outcomes: ${stdout}${stderr}`);
        return JSON.parse(line.slice(RESULT.length)) as ProtectionOutcomes;
      } finally { clearTimeout(deadline); }
    })();
    outcomes.set(engine, result);
  }
  return result;
}

for (const engine of ['wasm', 'native'] as const) {
  const reason = engine === 'native' ? nativeUnavailable : undefined;
  test.skipIf(reason !== undefined)(`protected cell input on ${engine}${reason ? ` (skipped: ${reason})` : ''}`, async () => {
    if (process.env[CHILD] === engine) {
      process.stdout.write(`${RESULT}${JSON.stringify(await protectedCellOutcomes(engine))}\n`);
    } else await engineOutcomes(engine);
  }, 30_000);
}

test.skipIf(nativeUnavailable !== undefined)(`protected cell input engine outcomes match${nativeUnavailable ? ` (skipped: ${nativeUnavailable})` : ''}`, async () => {
  expect(await engineOutcomes('native')).toEqual(await engineOutcomes('wasm'));
}, 60_000);

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
